// Azure OpenAI call: one clause in, one validated-shape JSON object out.
import { getConfig } from "./config.js";

const SYSTEM_PROMPT = `You are a contract clause analyst. You receive ONE clause from our company's own contract template, plus lists of allowed values. Draft our negotiation standard for that clause.

Rules
1. "category", "department" and "risk" MUST be copied exactly from the allowed lists. If none fits well, pick the closest and set "confidence" below 0.5.
2. "department" is the team that must review changes to this clause. Decide from the clause wording (e.g. invoices, fees, payment -> Finance; personal data, security, systems -> Security / IT; liability, indemnity, law, IP -> Legal).
3. The template is drafted in OUR favour. "ourRole" says which party we are (for example Customer or Supplier); if it is empty, infer it from the clause.
   "preferred": the position in the template wording, in one or two plain sentences.
   "fallback": a moderate concession TO THE OTHER PARTY that Legal could accept without escalation. It must be worse for us than "preferred", never better.
   "redLine": the point where a counterparty change becomes unacceptable for us and must be escalated or rejected. It describes a change in the other party's favour, never one in ours.
4. Figures (days, months, amounts, percentages, multiples) MUST be copied from the clause. Never introduce a number that is not in the clause. If a compromise needs a different number, describe it in words (for example "a longer payment period approved by Finance").
5. "keywords": 5 to 10 lowercase terms or short phrases that appear in the clause or are direct synonyms. They are used to match counterparty clauses to this one.
6. "risk": how serious it is for us if this clause is lost or weakened.
7. "title": a short clause name (2 to 5 words). Use the clause heading when one is given.
8. "confidence": 0 to 1. Use below 0.5 when the text is ambiguous, very long, covers several topics, or the category/department fit is poor.
9. Use only the text given. Do not use outside knowledge or describe other clauses.
10. If the text is not a contract clause (a heading, address block, table of contents, signature block or marketing text), set "skip": true and fill the other fields with empty values.`;

function schema(lists) {
  const str = { type: "string" };
  return {
    name: "clause_standard",
    strict: true,
    schema: {
      type: "object",
      additionalProperties: false,
      required: ["skip", "title", "category", "department", "risk", "preferred", "fallback", "redLine", "keywords", "confidence"],
      properties: {
        skip: { type: "boolean" },
        title: str,
        category: { type: "string", enum: lists.categories },
        department: { type: "string", enum: lists.departments },
        risk: { type: "string", enum: lists.risks },
        preferred: str,
        fallback: str,
        redLine: str,
        keywords: { type: "array", items: str },
        confidence: { type: "number" },
      },
    },
  };
}

// One chat call with strict JSON output. Retries on throttling and server errors.
export async function chatJson(system, user, jsonSchema, maxTokens, deployment) {
  const c = getConfig();
  const url = `${c.aoaiEndpoint}/openai/deployments/${deployment || c.aoaiDeployment}/chat/completions?api-version=${c.aoaiApiVersion}`;
  const body = JSON.stringify({
    temperature: 0,
    max_tokens: maxTokens,
    messages: [
      { role: "system", content: system },
      { role: "user", content: JSON.stringify(user) },
    ],
    response_format: { type: "json_schema", json_schema: jsonSchema },
  });

  for (let attempt = 0; ; attempt++) {
    let res;
    try {
      res = await fetch(url, { method: "POST", headers: { "api-key": c.aoaiKey, "Content-Type": "application/json" }, body });
    } catch (e) {
      if (attempt >= 4) throw e; // network drop: retry like a 5xx
      await new Promise((r) => setTimeout(r, 2 ** (attempt + 1) * 1000));
      continue;
    }
    if ((res.status === 429 || res.status >= 500) && attempt < 4) {
      const wait = Number(res.headers.get("retry-after")) || 2 ** (attempt + 1);
      await new Promise((r) => setTimeout(r, wait * 1000));
      continue;
    }
    const json = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(`Azure OpenAI ${res.status}: ${json.error?.code || "error"}`);
    const choice = json.choices?.[0];
    const usage = { input: json.usage?.prompt_tokens || 0, output: json.usage?.completion_tokens || 0 };
    if (choice?.finish_reason !== "stop") return { answer: null, usage, error: `AI stopped early (${choice?.finish_reason})` };
    try {
      return { answer: JSON.parse(choice.message.content), usage };
    } catch {
      return { answer: null, usage, error: "AI returned invalid JSON" };
    }
  }
}

export function draftClause(clause, ctx) {
  const user = {
    contractType: ctx.contractType,
    ourRole: ctx.ourRole || "",
    clauseHeading: clause.heading,
    sectionRef: clause.sectionRef,
    clauseText: clause.text,
    allowedCategories: ctx.lists.categories,
    allowedDepartments: ctx.lists.departments,
    allowedRiskLevels: ctx.lists.risks,
  };
  return chatJson(SYSTEM_PROMPT, user, schema(ctx.lists), 700);
}
