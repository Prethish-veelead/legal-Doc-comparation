// Part A: document link in, clause standard metadata out. The document is held in memory only.
import { createHash, randomUUID } from "node:crypto";
import { getConfig } from "./config.js";
import { downloadDocument, UserError } from "./graph.js";
import { extractBlocks } from "./extract.js";
import { splitClauses } from "./split.js";
import { draftClause } from "./ai.js";
import { checkAnswer, fallbackDraft } from "./validate.js";
import * as sp from "./sharepoint.js";

export const CONTRACT_TYPES = ["Master services", "NDA", "Data processing", "Statement of work", "Services", "SaaS subscription", "Software licence", "Renewal", "Reseller", "Data licence", "Consulting", "Other"];

// gpt-4o-mini Global Standard list prices per 1M tokens (override with settings if they change).
const PRICE_IN = Number(process.env.PRICE_INPUT_PER_M || 0.15);
const PRICE_OUT = Number(process.env.PRICE_OUTPUT_PER_M || 0.6);

export function validateRequest(body) {
  if (!body || typeof body !== "object") throw new UserError("Send a JSON body with documentUrl and contractType.");
  const allowed = new Set(["documentUrl", "contractType", "ourRole", "force"]);
  const extra = Object.keys(body).filter((k) => !allowed.has(k));
  if (extra.length) throw new UserError(`Unexpected field(s): ${extra.join(", ")}`);
  const { documentUrl, contractType, ourRole = "", force = false } = body;
  if (typeof documentUrl !== "string" || !documentUrl.trim() || documentUrl.length > 2000) throw new UserError("documentUrl is required (max 2000 characters).");
  if (!CONTRACT_TYPES.includes(contractType)) throw new UserError(`contractType must be one of: ${CONTRACT_TYPES.join(", ")}`);
  if (typeof ourRole !== "string" || ourRole.length > 40) throw new UserError("ourRole must be text of at most 40 characters, for example Customer or Supplier.");
  if (typeof force !== "boolean") throw new UserError("force must be true or false.");
  return { documentUrl: documentUrl.trim(), contractType, ourRole: ourRole.trim(), force };
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

export async function processDocument(input, log = () => {}) {
  const cfg = getConfig();
  const { documentUrl, contractType, ourRole, force } = validateRequest(input);
  const started = Date.now();
  const runId = `RUN-${new Date().toISOString().slice(0, 10).replace(/-/g, "")}-${randomUUID().slice(0, 8)}`;
  let runItemId = null;

  try {
    const doc = await downloadDocument(documentUrl);
    const contentHash = "sha256:" + createHash("sha256").update(doc.buffer).digest("hex");
    log(`Downloaded ${doc.fileName} (${doc.sizeBytes} bytes) into memory`);

    if (!force) {
      const previous = await sp.findCompletedRun(contentHash);
      if (previous) {
        log("Same file was processed before; returning the saved result");
        return { ...previous, reused: true, message: "This exact file was processed before, so the saved result is shown. To draft it again, tick \"Process again\" (API: force=true)." };
      }
    }

    runItemId = await sp.createRun({ Title: doc.fileName, LDC_RunId: runId, LDC_SourceUrl: { Url: documentUrl, Description: doc.fileName }, LDC_RunStatus: "Running", LDC_ContractType: contractType, LDC_ContentHash: contentHash, LDC_StartedOn: new Date(started).toISOString() }).catch(async (e) => {
      // Hyperlink fields can be picky; retry with the URL as plain text in the description only.
      if (e.status !== 400) throw e;
      return sp.createRun({ Title: doc.fileName, LDC_RunId: runId, LDC_RunStatus: "Running", LDC_ContractType: contractType, LDC_ContentHash: contentHash, LDC_StartedOn: new Date(started).toISOString() });
    });

    const { type, blocks, pages } = await extractBlocks(doc.fileName, doc.buffer);
    doc.buffer = null; // drop the file as soon as the text is out
    const split = splitClauses(blocks);
    log(`Found ${split.sectionsFound} sections, ${split.clauses.length} clauses to draft`);
    if (!split.clauses.length) throw new UserError("No clauses were found in this document.", 422);

    const lists = await sp.loadReferenceLists();
    const usage = { input: 0, output: 0, calls: 0 };
    const skippedByAi = [];

    const drafted = await mapLimit(split.clauses, cfg.aiConcurrency, async (clause, idx) => {
      let ai;
      try {
        ai = await draftClause(clause, { contractType, ourRole, lists });
      } catch (e) {
        log(`AI call failed for clause ${idx + 1}: ${e.message}`);
        return { ...clause, ...fallbackDraft(clause, lists, "the AI service did not respond"), draftedBy: "Fallback" };
      }
      usage.calls++;
      usage.input += ai.usage.input;
      usage.output += ai.usage.output;
      if (!ai.answer) return { ...clause, ...fallbackDraft(clause, lists, ai.error), draftedBy: "Fallback" };
      if (ai.answer.skip) {
        skippedByAi.push(`${clause.heading || clause.sectionRef || `Section ${idx + 1}`} (AI: not a clause)`);
        return null;
      }
      const checked = checkAnswer(ai.answer, clause, lists);
      if (!checked.valid) return { ...clause, ...fallbackDraft(clause, lists, `the AI answer failed checks: ${checked.problems.join("; ")}`), draftedBy: "Fallback" };
      return { ...clause, ...checked.draft, draftedBy: "AI" };
    });

    const clauses = drafted.filter(Boolean).map((c, i) => ({
      clauseCode: null,
      order: i + 1,
      sectionRef: c.sectionRef,
      title: c.title.trim(),
      category: c.category,
      department: c.department,
      risk: c.risk,
      preferred: c.preferred.trim(),
      fallback: (c.fallback || "").trim(),
      redLine: (c.redLine || "").trim(),
      keywords: c.keywords,
      confidence: c.confidence,
      needsReview: c.needsReview,
      reviewReasons: c.reviewReasons,
      draftedBy: c.draftedBy,
      status: "Draft",
    }));
    if (!clauses.length) throw new UserError("The AI found no contract clauses in this document.", 422);
    sp.assignCodes(clauses);

    const templateItemId = await sp.templateListItemId(doc.driveId, doc.itemId);
    const title = doc.fileName.replace(/\.[^.]+$/, "");
    const standardId = await sp.createStandard({ title, contractType, templateItemId, clauses, sourceFile: doc.fileName });
    await sp.saveClauses(standardId, clauses, lists);
    log(`Saved standard ${standardId} with ${clauses.length} clauses`);

    const tokens = usage.input + usage.output;
    const cost = Math.round(((usage.input * PRICE_IN + usage.output * PRICE_OUT) / 1e6) * 10000) / 10000;
    const skipped = [...split.skipped, ...skippedByAi];
    const result = {
      run: { runId, status: "Completed", startedOn: new Date(started).toISOString(), durationSec: Math.round((Date.now() - started) / 100) / 10, processedBy: "Azure Function (Flex Consumption)", model: cfg.aoaiDeployment, aiCalls: usage.calls, tokensUsed: tokens, estimatedCostUsd: cost },
      document: { fileName: doc.fileName, sourceUrl: documentUrl, fileType: type, fileSizeKb: Math.round(doc.sizeBytes / 1024), pages, contentHash, documentStored: false, textStored: false },
      standard: { standardId, title, contractType, version: "v1", status: "Draft" },
      summary: {
        sectionsFound: split.sectionsFound,
        skipped: skipped.length,
        skippedReasons: skipped,
        clausesDrafted: clauses.length,
        aiDrafted: clauses.filter((c) => c.draftedBy === "AI").length,
        fallbackUsed: clauses.filter((c) => c.draftedBy === "Fallback").length,
        needsReview: clauses.filter((c) => c.needsReview).length,
      },
      clauses,
    };

    await sp.updateRun(runItemId, {
      LDC_RunStatus: "Completed",
      LDC_StandardLookupId: standardId,
      LDC_FileType: type,
      LDC_FileSizeKb: result.document.fileSizeKb,
      LDC_SectionsFound: split.sectionsFound,
      LDC_ClauseCount: clauses.length,
      LDC_AiDrafted: result.summary.aiDrafted,
      LDC_FallbackUsed: result.summary.fallbackUsed,
      LDC_NeedsReviewCount: result.summary.needsReview,
      LDC_Model: cfg.aoaiDeployment,
      LDC_TokensUsed: tokens,
      LDC_EstimatedCostUsd: cost,
      LDC_DurationSec: result.run.durationSec,
      LDC_ResultJson: JSON.stringify(result),
    });
    return result;
  } catch (e) {
    if (runItemId) {
      await sp
        .updateRun(runItemId, { LDC_RunStatus: "Failed", LDC_Error: e instanceof UserError ? e.message : "Processing failed. See the function logs.", LDC_DurationSec: Math.round((Date.now() - started) / 1000) })
        .catch(() => {});
    }
    throw e;
  }
}
