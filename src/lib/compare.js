// Part B: the other party's draft in, findings against a published clause standard out.
// The document is held in memory only. Nothing is written to SharePoint here.
import { createHash, randomUUID } from "node:crypto";
import { getConfig } from "./config.js";
import { downloadDocument, UserError } from "./graph.js";
import { extractBlocks } from "./extract.js";
import { splitClauses } from "./split.js";
import { chatJson } from "./ai.js";
import * as sp from "./sharepoint.js";

const PROMPT_VERSION = "compare-1";
// gpt-4.1-mini Global Standard list prices per 1M tokens (override with settings if they change).
const PRICE_IN = Number(process.env.COMPARE_PRICE_INPUT_PER_M || 0.4);
const PRICE_OUT = Number(process.env.COMPARE_PRICE_OUTPUT_PER_M || 1.6);
const CANDIDATES = 3;

export const VERDICTS = ["meets_preferred", "within_fallback", "outside_fallback", "crosses_red_line", "missing", "unclear"];

const SYSTEM_PROMPT = `You check the other party's contract draft against ONE rule from our contract standard.
You receive the rule (preferred, fallback and red-line positions) and up to ${CANDIDATES} candidate clauses from their draft, each with an id.

Rules
1. "matchedClause": the id of the candidate that covers the same topic as the rule. If none covers it, use "none" and set verdict "missing".
2. "analysis" (write this BEFORE choosing the verdict): list the key terms and figures in their clause, then test them in this order:
   a) Does any term meet the red-line description? For each figure write "value vs limit", for example "20 days vs under 60 days: 20 < 60, met" or "3% vs above 3%: 3 > 3 is false, not met".
      A value EQUAL to an "under", "below", "less than", "above", "over" or "more than" limit does NOT meet it.
   b) If not, is every term at least as good for us as the fallback? A term equal to the fallback figure is within the fallback.
   c) Is every term at least as good for us as the preferred position?
   Keep it short (at most 80 words).
3. "verdict", judged from OUR side ("ourRole" says which party we are; if empty, infer it from the rule). It MUST agree with the analysis:
   crosses_red_line: test (a) is met by any term. This wins over every other verdict.
   meets_preferred: the clause gives us the preferred position or something better for us.
   within_fallback: worse for us than preferred, but no worse than the fallback position.
   outside_fallback: worse than the fallback, but no term meets the red-line description.
   missing: no candidate covers the rule's topic.
   unclear: a candidate covers the topic but the text is too vague or incomplete to judge.
4. Compare figures carefully (days, months, percentages, amounts, notice periods). Words and digits mean the same ("fifteen (15) days" = 15 days).
   If the clause states the same term more than once with different values (for example two different payment periods), judge by the value that is WORST for us, quote that sentence, say in the explanation that the clause contains conflicting terms, and set confidence below 0.6.
5. Never mention candidate ids (C1, C2, ...) in "analysis" or "explanation"; name the clause by its heading.
6. "quote": copy the exact words (at most 40 words) from the matched clause that decide the verdict. Copy them character for character. Empty when the verdict is missing.
7. "explanation": one plain-English sentence saying what their clause says and how it compares with our standard. Only use figures that appear in the clause or the rule.
8. "confidence": 0 to 1. Use below 0.6 when the match or the verdict is uncertain.
9. Use only the text given.`;

function schema(candidateIds) {
  return {
    name: "rule_check",
    strict: true,
    schema: {
      type: "object",
      additionalProperties: false,
      required: ["matchedClause", "analysis", "verdict", "quote", "explanation", "confidence"],
      properties: {
        matchedClause: { type: "string", enum: ["none", ...candidateIds] },
        analysis: { type: "string" },
        verdict: { type: "string", enum: VERDICTS },
        quote: { type: "string" },
        explanation: { type: "string" },
        confidence: { type: "number" },
      },
    },
  };
}

export function validateCompareRequest(body) {
  if (!body || typeof body !== "object") throw new UserError("Send a JSON body with documentUrl and standardId.");
  const allowed = new Set(["documentUrl", "standardId", "ourRole"]);
  const extra = Object.keys(body).filter((k) => !allowed.has(k));
  if (extra.length) throw new UserError(`Unexpected field(s): ${extra.join(", ")}`);
  const { documentUrl, standardId, ourRole = "" } = body;
  if (typeof documentUrl !== "string" || !documentUrl.trim() || documentUrl.length > 2000) throw new UserError("documentUrl is required (max 2000 characters).");
  if (!Number.isInteger(standardId) || standardId <= 0) throw new UserError("standardId must be a positive whole number.");
  if (typeof ourRole !== "string" || ourRole.length > 40) throw new UserError("ourRole must be text of at most 40 characters, for example Customer or Supplier.");
  return { documentUrl: documentUrl.trim(), standardId, ourRole: ourRole.trim() };
}

// ---------- matching: which of their clauses could cover a rule ----------
const STOP = new Set("the and for with that this shall will may any all its not are from such other under than into each upon only per our their party parties agreement days".split(" "));
const words = (t) => (t || "").toLowerCase().match(/[a-z]{3,}/g)?.filter((w) => !STOP.has(w)) || [];

function candidatesFor(rule, clauses) {
  const ruleWords = new Map();
  for (const w of words(`${rule.preferred} ${rule.fallback} ${rule.redLine}`)) ruleWords.set(w, 1);
  for (const w of words(rule.title)) ruleWords.set(w, 3);
  return clauses
    .map((c) => {
      const head = new Set(words(c.heading));
      const body = new Set(words(c.text));
      let score = 0;
      for (const [w, weight] of ruleWords) score += head.has(w) ? weight * 2 : body.has(w) ? weight : 0;
      return { c, score };
    })
    .filter((x) => x.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, CANDIDATES)
    .map((x) => x.c);
}

// ---------- checks on the AI answer ----------
const norm = (t) =>
  (t || "")
    .toLowerCase()
    .replace(/[‘’“”]/g, "'")
    .replace(/["']/g, "'")
    .replace(/\s+/g, " ")
    .trim();

function checkVerdict(answer, byId) {
  const reasons = [];
  let { verdict, matchedClause } = answer;
  let confidence = Math.max(0, Math.min(1, Number(answer.confidence) || 0));
  const clause = byId.get(matchedClause) || null;
  if (verdict !== "missing" && !clause) {
    verdict = "unclear";
    reasons.push("The AI gave a verdict without naming a matching clause");
  }
  if (verdict === "missing" && clause) matchedClause = "none";
  if (clause && verdict !== "missing") {
    if (!answer.quote.trim()) reasons.push("The AI gave no quote from the clause");
    else if (!norm(clause.text).includes(norm(answer.quote)) && !norm(clause.heading).includes(norm(answer.quote))) {
      reasons.push("The AI's quote does not appear word for word in the clause");
      confidence = Math.min(confidence, 0.45);
    }
  }
  // Small models sometimes call a value EQUAL to an "above/under" limit a red line. Their reasoning
  // usually shows it ("1.5 = 1.5", "is false", "does not meet"), so send those cases to a person.
  const doubtful = /\bno red[- ]line\b|\bis false\b|\bdoes not meet\b|\bnot (above|under|below|over|more than|less than)\b|(\d+(?:\.\d+)?)\s*=\s*\2(?!\d)/i;
  if (verdict === "crosses_red_line" && doubtful.test(answer.analysis || "")) {
    reasons.push("The AI's reasoning suggests a limit may not be crossed (a value equals the limit); please confirm the red line");
    confidence = Math.min(confidence, 0.45);
  }
  if (confidence < 0.6) reasons.push("AI confidence below 0.6");
  return { verdict, clause: verdict === "missing" ? null : clause, confidence, reasons };
}

// ---------- routing to a review team ----------
function classify(text, rules, lists) {
  const t = (text || "").toLowerCase();
  const score = {};
  const hits = {};
  for (const r of rules) {
    if (r.exclude.some((x) => t.includes(x))) continue;
    const matched = r.patterns.filter((p) => new RegExp(`\\b${p.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`).test(t));
    if (!matched.length) continue;
    score[r.categoryKey] = (score[r.categoryKey] || 0) + matched.length * r.weight;
    (hits[r.categoryKey] ||= []).push(r.title);
  }
  const best = Object.entries(score).sort((a, b) => b[1] - a[1])[0];
  const titleFor = Object.fromEntries(Object.entries(lists.departmentKeys).map(([title, key]) => [key, title]));
  if (best && titleFor[best[0]]) return { team: titleFor[best[0]], note: `Rule '${hits[best[0]].join("', '")}' → ${titleFor[best[0]]}` };
  return { team: lists.fallbackDepartment, note: `No rule matched → ${lists.fallbackDepartment}` };
}

function routeRule(rule, clause, classRules, lists) {
  const team = (rule.signOffTeams || []).find((t) => lists.departments.includes(t));
  if (team) return { team, note: `Contract standard '${rule.title}' → ${team}` };
  return classify(`${rule.title} ${clause?.heading || ""} ${clause?.text || ""}`, classRules, lists);
}

// ---------- findings in the existing app's LDC_FindingsJson format ----------
const SEVERITY = (verdict, risk) =>
  ({ crosses_red_line: risk, missing: risk, outside_fallback: risk === "low" ? "low" : "medium", unclear: "medium", within_fallback: "low" })[verdict];
const TITLE = { crosses_red_line: "red line crossed", outside_fallback: "outside fallback", within_fallback: "within fallback", missing: "clause missing", unclear: "needs a closer look" };
const DESCRIPTION = {
  crosses_red_line: "Their wording crosses our red line for this clause.",
  outside_fallback: "Their wording is worse for us than our fallback position.",
  within_fallback: "Their wording is acceptable under our fallback position.",
  missing: "Their draft has no clause covering this rule.",
  unclear: "A clause covers this topic, but it could not be judged with confidence.",
};

function toFinding(n, rule, check, answer, route, lists) {
  const why = check.verdict === "crosses_red_line" ? `Red line: ${rule.redLine}` : check.verdict === "within_fallback" ? `Fallback: ${rule.fallback}` : `The contract standard requires: ${rule.preferred}`;
  return {
    id: `F${n}`,
    clause: check.clause ? `§${check.clause.sectionRef || check.clause.heading || "?"}` : "—",
    target: check.clause?.sectionRef || "",
    title: `${rule.title}: ${TITLE[check.verdict]}`,
    severity: SEVERITY(check.verdict, rule.risk || "medium"),
    ruleCode: rule.code,
    category: rule.category,
    redLine: check.verdict === "crosses_red_line",
    confidence: Math.round(check.confidence * 100) / 100,
    proposed: check.clause ? answer.quote.trim() : "",
    required: rule.preferred,
    description: answer.explanation?.trim() || DESCRIPTION[check.verdict],
    aiReasoning: answer.analysis?.trim() || "",
    why,
    routedTeam: route.team,
    categoryKey: lists.departmentKeys[route.team] || null,
    routeNote: route.note,
    verdict: check.verdict,
    needsReview: check.reasons.length > 0,
    reviewReasons: check.reasons,
    activity: [],
  };
}

function summaryText(standard, findings, notInStandard) {
  const material = findings.filter((f) => f.verdict !== "within_fallback");
  if (!material.length) return `No material findings against ${`${standard.title} ${standard.version}`.trim()}.${notInStandard.length ? ` ${notInStandard.length} clause(s) are not covered by the standard.` : ""}`;
  const count = (s) => material.filter((f) => f.severity === s).length;
  const red = findings.filter((f) => f.redLine).map((f) => f.title.split(":")[0]);
  const top = [...material].sort((a, b) => "hml".indexOf(a.severity[0]) - "hml".indexOf(b.severity[0]))[0];
  return [
    `${material.length} material finding(s) against ${`${standard.title} ${standard.version}`.trim()}: ${count("high")} high, ${count("medium")} medium, ${count("low")} low.`,
    red.length ? `Red lines crossed: ${red.join(", ")}.` : "No red lines crossed.",
    `Top risk: ${top.title}.`,
    notInStandard.length ? `${notInStandard.length} clause(s) are not covered by the standard.` : "",
  ]
    .filter(Boolean)
    .join(" ");
}

// Same formula and bands as the existing SPFx app, so scores match its Analysis panel and approval thresholds.
const SEVERITY_WEIGHT = { critical: 30, high: 20, medium: 10, low: 4 };
function riskOf(findings) {
  const points = findings.reduce((s, f) => s + (SEVERITY_WEIGHT[f.severity] || 0) + (f.redLine ? 10 : 0), 0);
  const riskScore = Math.round(100 * (1 - Math.exp(-points / 70)));
  const riskLevel = riskScore >= 85 ? "Critical" : riskScore >= 65 ? "High" : riskScore >= 40 ? "Medium" : "Low";
  return { riskScore, riskLevel };
}

async function mapLimit(items, limit, fn) {
  const out = new Array(items.length);
  let i = 0;
  await Promise.all(
    Array.from({ length: Math.min(limit, items.length) }, async () => {
      while (i < items.length) {
        const idx = i++;
        out[idx] = await fn(items[idx], idx);
      }
    }),
  );
  return out;
}

// `localDoc` ({ fileName, buffer }) is only passed by the command-line test runner.
export async function compareDocument(input, log = () => {}, localDoc = null) {
  const cfg = getConfig();
  const { documentUrl, standardId, ourRole } = validateCompareRequest(input);
  const started = Date.now();
  const runId = `CMP-${new Date().toISOString().slice(0, 10).replace(/-/g, "")}-${randomUUID().slice(0, 8)}`;

  const standard = await sp.getPublishedStandard(standardId);
  if (!standard) throw new UserError("No contract standard was found with that ID.", 404);
  if (standard.status !== "Published" || !standard.rules.length) throw new UserError(`Standard ${standardId} (${standard.title}) is not published yet. Publish it first.`, 409);

  const doc = localDoc ? { ...localDoc, sizeBytes: localDoc.buffer.length } : await downloadDocument(documentUrl);
  const contentHash = "sha256:" + createHash("sha256").update(doc.buffer).digest("hex");
  log(`Downloaded ${doc.fileName} (${doc.sizeBytes} bytes) into memory`);
  const { type, blocks, pages } = await extractBlocks(doc.fileName, doc.buffer);
  doc.buffer = null; // drop the file as soon as the text is out
  const split = splitClauses(blocks);
  if (!split.clauses.length) throw new UserError("No clauses were found in this document.", 422);
  const clauses = split.clauses.map((c, i) => ({ ...c, id: `C${i + 1}` }));
  log(`Found ${clauses.length} clauses; checking ${standard.rules.length} rules`);

  const [lists, classRules] = await Promise.all([sp.loadReferenceLists(), sp.loadClassificationRules()]);
  const usage = { input: 0, output: 0, calls: 0 };
  const byId = new Map(clauses.map((c) => [c.id, c]));
  const rules = [...standard.rules].sort((a, b) => (a.order ?? 99) - (b.order ?? 99));

  const checks = await mapLimit(rules, cfg.aiConcurrency, async (rule) => {
    const cands = candidatesFor(rule, clauses);
    let ai;
    try {
      ai = await chatJson(
        SYSTEM_PROMPT,
        {
          ourRole,
          contractType: standard.contractType,
          rule: { title: rule.title, preferred: rule.preferred, fallback: rule.fallback, redLine: rule.redLine },
          candidates: cands.map((c) => ({ id: c.id, heading: c.heading, text: c.text })),
        },
        schema(cands.map((c) => c.id)),
        700,
        cfg.compareDeployment,
      );
    } catch (e) {
      log(`AI call failed for rule ${rule.code}: ${e.message}`);
      ai = { answer: null, usage: { input: 0, output: 0 }, error: "the AI service did not respond" };
    }
    usage.calls++;
    usage.input += ai.usage.input;
    usage.output += ai.usage.output;
    if (!ai.answer) {
      const best = cands[0] || null;
      return { rule, answer: { quote: "", explanation: "", analysis: "" }, check: { verdict: best ? "unclear" : "missing", clause: best, confidence: 0, reasons: [`No AI verdict: ${ai.error}`] } };
    }
    return { rule, answer: ai.answer, check: checkVerdict(ai.answer, byId) };
  });

  let n = 0;
  const findings = [];
  for (const { rule, answer, check } of checks) {
    if (check.verdict === "meets_preferred" && !check.reasons.length) continue;
    findings.push(toFinding(++n, rule, check, answer, routeRule(rule, check.clause, classRules, lists), lists));
  }

  const used = new Set(checks.map((x) => x.check.clause?.id).filter(Boolean));
  const numbered = clauses.some((c) => c.sectionRef);
  const notInStandard = clauses
    .filter((c) => !used.has(c.id) && (!numbered || c.sectionRef)) // in a numbered document, unnumbered blocks are titles or notes
    .map((c) => {
      const route = classify(`${c.heading || ""} ${c.text}`, classRules, lists);
      return { clause: `§${c.sectionRef || c.heading || c.id}`, title: c.heading || c.text.slice(0, 60), suggestedTeam: route.team, routeNote: route.note };
    });

  const durationSec = Math.round((Date.now() - started) / 100) / 10;
  const tokens = usage.input + usage.output;
  const cost = Math.round(((usage.input * PRICE_IN + usage.output * PRICE_OUT) / 1e6) * 10000) / 10000;
  const count = (v) => checks.filter((x) => x.check.verdict === v).length;
  const aiSummary = summaryText(standard, findings, notInStandard);

  return {
    run: { runId, status: "Completed", startedOn: new Date(started).toISOString(), durationSec, model: cfg.compareDeployment, aiCalls: usage.calls, tokensUsed: tokens, estimatedCostUsd: cost, savedToSharePoint: false },
    document: { fileName: doc.fileName, sourceUrl: documentUrl, fileType: type, fileSizeKb: Math.round(doc.sizeBytes / 1024), pages, contentHash, documentStored: false, textStored: false },
    standard: { standardId: standard.standardId, title: standard.title, version: standard.version, rulesChecked: rules.length },
    summary: {
      clausesFound: clauses.length,
      meetsPreferred: count("meets_preferred"),
      withinFallback: count("within_fallback"),
      outsideFallback: count("outside_fallback"),
      redLinesCrossed: count("crosses_red_line"),
      missing: count("missing"),
      unclear: count("unclear"),
      needsReview: findings.filter((f) => f.needsReview).length,
      notInStandard: notInStandard.length,
      ...riskOf(findings),
      aiSummary,
    },
    checks: checks.map(({ rule, check }) => ({ ruleCode: rule.code, rule: rule.title, verdict: check.verdict, clause: check.clause ? `§${check.clause.sectionRef || check.clause.heading}` : null, confidence: Math.round(check.confidence * 100) / 100 })),
    // Same shape as LDC_Comparisons.LDC_FindingsJson, so the existing app can show it.
    findingsJson: {
      schemaVersion: 1,
      analysis: { completedOn: new Date().toISOString(), durationSeconds: Math.round(durationSec), confidence: findings.length ? Math.round((findings.reduce((s, f) => s + f.confidence, 0) / findings.length) * 100) / 100 : 1, totalChanges: findings.length, materialChanges: findings.filter((f) => f.verdict !== "within_fallback").length, model: cfg.compareDeployment, promptVersion: PROMPT_VERSION, diffEngineVersion: "n/a (AI rule check)", correlationId: runId },
      findings,
    },
    notInStandard,
  };
}
