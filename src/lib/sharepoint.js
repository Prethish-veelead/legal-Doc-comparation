// SharePoint list access: reference data, list setup and saving metadata.
import { graph, getSiteId } from "./graph.js";

export const DEFAULT_CATEGORIES = ["Commercial", "Liability", "Term", "Privacy", "Confidentiality", "Intellectual property", "Warranty", "Assignment", "Employment", "Delivery", "General"];
export const RISKS = ["low", "medium", "high"];
const CATEGORY_SETTING = "clauseCategories";

const multi = (name) => ({ name, text: { allowMultipleLines: true, linesForEditing: 6, textType: "plain" } });
const text = (name) => ({ name, text: {} });
const num = (name) => ({ name, number: {} });
const bool = (name) => ({ name, boolean: {} });
const choice = (name, choices) => ({ name, choice: { choices, displayAs: "dropDownMenu" } });

// New lists owned by Part A. `lookupTo` columns are resolved at setup time.
export const LIST_DEFS = {
  LDC_Clauses: {
    displayName: "Clauses",
    columns: [
      { name: "LDC_Standard", lookupTo: "LDC_ContractStandards" },
      text("LDC_ClauseCode"),
      text("LDC_SectionRef"),
      num("LDC_Order"),
      text("LDC_Category"),
      text("LDC_Department"),
      choice("LDC_Risk", RISKS),
      multi("LDC_Preferred"),
      multi("LDC_Fallback"),
      multi("LDC_RedLine"),
      multi("LDC_Keywords"),
      num("LDC_Confidence"),
      bool("LDC_NeedsReview"),
      multi("LDC_ReviewReasons"),
      choice("LDC_DraftedBy", ["AI", "Fallback", "Person"]),
      choice("LDC_ClauseStatus", ["Draft", "Published"]),
    ],
  },
  LDC_ProcessingRuns: {
    displayName: "Processing Runs",
    columns: [
      text("LDC_RunId"),
      { name: "LDC_SourceUrl", hyperlinkOrPicture: { isPicture: false } },
      text("LDC_FileType"),
      num("LDC_FileSizeKb"),
      text("LDC_ContentHash"),
      { name: "LDC_Standard", lookupTo: "LDC_ContractStandards" },
      text("LDC_ContractType"),
      choice("LDC_RunStatus", ["Running", "Completed", "Failed"]),
      num("LDC_SectionsFound"),
      num("LDC_ClauseCount"),
      num("LDC_AiDrafted"),
      num("LDC_FallbackUsed"),
      num("LDC_NeedsReviewCount"),
      text("LDC_Model"),
      num("LDC_TokensUsed"),
      num("LDC_EstimatedCostUsd"),
      { name: "LDC_StartedOn", dateTime: { format: "dateTime" } },
      num("LDC_DurationSec"),
      multi("LDC_Error"),
      multi("LDC_ResultJson"),
    ],
  },
};

let listCache = null;
async function lists() {
  if (listCache) return listCache;
  const sid = await getSiteId();
  const all = (await graph(`/sites/${sid}/lists?$select=id,webUrl,displayName&$top=200`)).value;
  listCache = Object.fromEntries(all.map((l) => [decodeURIComponent(l.webUrl.split("/").pop()), l.id]));
  return listCache;
}
export async function listId(urlName) {
  const id = (await lists())[urlName];
  if (!id) throw new Error(`SharePoint list ${urlName} was not found. Run "npm run setup-lists" first.`);
  return id;
}
const itemsPath = async (urlName) => `/sites/${await getSiteId()}/lists/${await listId(urlName)}/items`;

async function allItems(urlName, filter) {
  let url = `${await itemsPath(urlName)}?$expand=fields&$top=500${filter ? `&$filter=${encodeURIComponent(filter)}` : ""}`;
  const out = [];
  while (url) {
    const page = await graph(url, { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } });
    out.push(...page.value);
    url = page["@odata.nextLink"];
  }
  return out;
}

// ---------- setup ----------
export async function ensureLists(log = console.log) {
  const sid = await getSiteId();
  for (const [urlName, def] of Object.entries(LIST_DEFS)) {
    const cols = await Promise.all(
      def.columns.map(async (c) => (c.lookupTo ? { name: c.name, lookup: { listId: await listId(c.lookupTo), columnName: "Title" } } : c)),
    );
    let id = (await lists())[urlName];
    if (!id) {
      const created = await graph(`/sites/${sid}/lists`, { method: "POST", body: { displayName: urlName, columns: cols, list: { template: "genericList" } } });
      id = created.id;
      await graph(`/sites/${sid}/lists/${id}`, { method: "PATCH", body: { displayName: def.displayName } });
      listCache = null;
      log(`Created list ${def.displayName} (${urlName}) with ${cols.length} columns`);
      continue;
    }
    const existing = new Set((await graph(`/sites/${sid}/lists/${id}/columns?$select=name`)).value.map((c) => c.name));
    const missing = cols.filter((c) => !existing.has(c.name));
    for (const c of missing) await graph(`/sites/${sid}/lists/${id}/columns`, { method: "POST", body: c });
    log(`List ${def.displayName} (${urlName}) exists; added ${missing.length} missing column(s)`);
  }

  const settings = await allItems("LDC_AppSettings");
  if (!settings.some((s) => s.fields.Title === CATEGORY_SETTING)) {
    await graph(await itemsPath("LDC_AppSettings"), {
      method: "POST",
      body: { fields: { Title: CATEGORY_SETTING, LDC_SettingValue: JSON.stringify({ schemaVersion: 1, categories: DEFAULT_CATEGORIES }), LDC_SettingVersion: 1, LDC_Description: "Clause categories the AI may choose from when drafting clause standards." } },
    });
    log(`Added ${CATEGORY_SETTING} to App Settings`);
  }
}

// ---------- reference data ----------
export async function loadReferenceLists() {
  const [cats, settings] = await Promise.all([allItems("LDC_ReviewCategories"), allItems("LDC_AppSettings")]);
  const active = cats.map((i) => i.fields).filter((f) => f.LDC_IsActive !== false).sort((a, b) => (a.LDC_Order ?? 99) - (b.LDC_Order ?? 99));
  const fallback = active.find((f) => f.LDC_IsFallback) || active[active.length - 1];
  let categories = DEFAULT_CATEGORIES;
  const s = settings.find((x) => x.fields.Title === CATEGORY_SETTING);
  if (s) {
    try {
      const parsed = JSON.parse(s.fields.LDC_SettingValue);
      if (Array.isArray(parsed.categories) && parsed.categories.length) categories = parsed.categories;
    } catch {
      /* keep defaults */
    }
  }
  return {
    departments: active.map((f) => f.Title),
    departmentKeys: Object.fromEntries(active.map((f) => [f.Title, f.LDC_CategoryKey])),
    fallbackDepartment: fallback.Title,
    categories,
    risks: RISKS,
  };
}

// ---------- runs ----------
export async function findCompletedRun(contentHash) {
  const rows = await allItems("LDC_ProcessingRuns", `fields/LDC_ContentHash eq '${contentHash}' and fields/LDC_RunStatus eq 'Completed'`);
  const latest = rows.sort((a, b) => Number(b.id) - Number(a.id))[0];
  return latest?.fields.LDC_ResultJson ? JSON.parse(latest.fields.LDC_ResultJson) : null;
}

export async function createRun(fields) {
  return (await graph(await itemsPath("LDC_ProcessingRuns"), { method: "POST", body: { fields } })).id;
}
export async function updateRun(itemId, fields) {
  await graph(`${await itemsPath("LDC_ProcessingRuns")}/${itemId}/fields`, { method: "PATCH", body: fields });
}

// ---------- standards and clauses ----------
const CODE_PREFIX = { Commercial: "COM", Liability: "LIAB", Term: "TERM", Privacy: "PRIV", Confidentiality: "CONF", "Intellectual property": "IP", Warranty: "WARR", Assignment: "ASSN", Employment: "EMP", Delivery: "DEL", General: "GEN" };
export function assignCodes(clauses) {
  const n = {};
  for (const c of clauses) {
    const p = CODE_PREFIX[c.category] || c.category.replace(/[^A-Za-z]/g, "").slice(0, 4).toUpperCase() || "GEN";
    n[p] = (n[p] || 0) + 1;
    c.clauseCode = `${p}-${String(n[p]).padStart(2, "0")}`;
  }
}

// Same shape the existing app stores in LDC_DraftClausesJson / LDC_ClausesJson.
export const toAppJson = (clauses) =>
  JSON.stringify(clauses.map((c) => ({ code: c.clauseCode, title: c.title, category: c.category, risk: c.risk, preferred: c.preferred, fallback: c.fallback, redLine: c.redLine, signOffTeams: [c.department], order: c.order })));

export async function templateListItemId(driveId, itemId) {
  try {
    return Number((await graph(`/drives/${driveId}/items/${itemId}/listItem?$select=id`)).id);
  } catch {
    return null;
  }
}

export async function createStandard({ title, contractType, templateItemId, clauses, sourceFile }) {
  const fields = {
    Title: title,
    LDC_ContractType: contractType,
    LDC_Version: "v1",
    LDC_StandardStatus: "Draft",
    LDC_OwnerTeam: "Legal",
    LDC_Description: `Drafted by AI from ${sourceFile} on ${new Date().toISOString().slice(0, 10)}. Review before publishing.`,
    LDC_ClauseCount: clauses.length,
    LDC_DraftClausesJson: toAppJson(clauses),
  };
  if (templateItemId) fields.LDC_TemplateDocumentLookupId = templateItemId;
  return Number((await graph(await itemsPath("LDC_ContractStandards"), { method: "POST", body: { fields } })).id);
}

export async function saveClauses(standardId, clauses, lists, concurrency = 4) {
  const path = await itemsPath("LDC_Clauses");
  let i = 0;
  const worker = async () => {
    while (i < clauses.length) {
      const c = clauses[i++];
      const created = await graph(path, {
        method: "POST",
        body: {
          fields: {
            Title: c.title,
            LDC_StandardLookupId: standardId,
            LDC_ClauseCode: c.clauseCode,
            LDC_SectionRef: c.sectionRef || "",
            LDC_Order: c.order,
            LDC_Category: c.category,
            LDC_Department: lists.departmentKeys[c.department] || c.department,
            LDC_Risk: c.risk,
            LDC_Preferred: c.preferred,
            LDC_Fallback: c.fallback,
            LDC_RedLine: c.redLine,
            LDC_Keywords: c.keywords.join(", "),
            LDC_Confidence: c.confidence,
            LDC_NeedsReview: c.needsReview,
            LDC_ReviewReasons: c.reviewReasons.join("\n"),
            LDC_DraftedBy: c.draftedBy,
            LDC_ClauseStatus: "Draft",
          },
        },
      });
      c.listItemId = Number(created.id);
    }
  };
  await Promise.all(Array.from({ length: Math.min(concurrency, clauses.length) }, worker));
}

// Publish: rebuild the rules from the Clauses list (so edits made there are kept), freeze them on the standard.
export async function publishStandard(standardId) {
  const std = await graph(`${await itemsPath("LDC_ContractStandards")}/${standardId}?$expand=fields`).catch((e) => {
    if (e.status === 404) return null;
    throw e;
  });
  if (!std) return null;
  const lists = await loadReferenceLists();
  const titleFor = Object.fromEntries(Object.entries(lists.departmentKeys).map(([t, k]) => [k, t]));
  const rows = (await allItems("LDC_Clauses", `fields/LDC_StandardLookupId eq ${Number(standardId)}`)).sort((a, b) => a.fields.LDC_Order - b.fields.LDC_Order);
  if (!rows.length) return { standardId, published: false, message: "This standard has no clauses in the Clauses list." };
  const clauses = rows.map((r) => ({
    clauseCode: r.fields.LDC_ClauseCode,
    title: r.fields.Title,
    category: r.fields.LDC_Category,
    risk: r.fields.LDC_Risk,
    preferred: r.fields.LDC_Preferred || "",
    fallback: r.fields.LDC_Fallback || "",
    redLine: r.fields.LDC_RedLine || "",
    department: titleFor[r.fields.LDC_Department] || r.fields.LDC_Department,
    order: r.fields.LDC_Order,
  }));
  await graph(`${await itemsPath("LDC_ContractStandards")}/${standardId}/fields`, {
    method: "PATCH",
    body: { LDC_ClausesJson: toAppJson(clauses), LDC_ClauseCount: clauses.length, LDC_StandardStatus: "Published", LDC_PublishedOn: new Date().toISOString() },
  });
  const cpath = await itemsPath("LDC_Clauses");
  for (const r of rows) await graph(`${cpath}/${r.id}/fields`, { method: "PATCH", body: { LDC_ClauseStatus: "Published" } });
  return { standardId: Number(standardId), title: std.fields.Title, published: true, clauseCount: clauses.length };
}

// ---------- Part B: reading a published standard and the routing rules ----------
export async function getPublishedStandard(standardId) {
  const item = await graph(`${await itemsPath("LDC_ContractStandards")}/${Number(standardId)}?$expand=fields`).catch((e) => {
    if (e.status === 404) return null;
    throw e;
  });
  if (!item) return null;
  const f = item.fields;
  let rules = [];
  try {
    rules = JSON.parse(f.LDC_ClausesJson || "[]");
  } catch {
    /* treated as not published below */
  }
  return {
    standardId: Number(item.id),
    title: f.Title,
    version: f.LDC_Version || "",
    contractType: f.LDC_ContractType || "",
    status: f.LDC_StandardStatus,
    rules: Array.isArray(rules) ? rules : [],
  };
}

export async function loadClassificationRules() {
  return (await allItems("LDC_ClassificationRules"))
    .map((i) => i.fields)
    .filter((f) => f.LDC_IsActive !== false && f.LDC_CategoryKey && f.LDC_Pattern)
    .map((f) => ({
      title: f.Title,
      categoryKey: f.LDC_CategoryKey,
      weight: Number(f.LDC_Weight) || 1,
      patterns: f.LDC_Pattern.split(/\r?\n/).map((p) => p.trim().toLowerCase()).filter(Boolean),
      exclude: (f.LDC_Exclude || "").split(/\r?\n/).map((p) => p.trim().toLowerCase()).filter(Boolean),
    }));
}
