// HTTP endpoints. All require the Function key (authLevel "function").
import { app } from "@azure/functions";
import { readFileSync } from "node:fs";
import { randomBytes } from "node:crypto";
import { processDocument, CONTRACT_TYPES } from "../lib/pipeline.js";
import { compareDocument } from "../lib/compare.js";
import { publishStandard } from "../lib/sharepoint.js";
import { UserError } from "../lib/graph.js";
import { getConfig } from "../lib/config.js";

const UI = readFileSync(new URL("../ui/index.html", import.meta.url), "utf8");
const BASE_HEADERS = {
  "X-Content-Type-Options": "nosniff",
  "X-Frame-Options": "DENY",
  "Referrer-Policy": "no-referrer",
  "Cache-Control": "no-store",
  "Strict-Transport-Security": "max-age=31536000",
};
const json = (status, body) => ({ status, headers: { ...BASE_HEADERS, "Content-Type": "application/json" }, body: JSON.stringify(body) });

async function readJson(request) {
  const len = Number(request.headers.get("content-length") || 0);
  if (len > 10_000) throw new UserError("Request body is too large.", 413);
  try {
    return await request.json();
  } catch {
    throw new UserError("Request body must be valid JSON.");
  }
}

function fail(context, e) {
  if (e instanceof UserError) return json(e.status, { error: e.message });
  context.error(`Unhandled error: ${e.message}`); // messages never contain document text
  return json(500, { error: "Processing failed. Please try again, or ask the administrator to check the function logs." });
}

app.http("process", {
  methods: ["POST"],
  authLevel: "function",
  route: "process",
  handler: async (request, context) => {
    try {
      const body = await readJson(request);
      const result = await processDocument(body, (m) => context.log(m));
      return json(200, result);
    } catch (e) {
      return fail(context, e);
    }
  },
});

// Part B: check the other party's draft against a published standard. Read-only: nothing is saved.
app.http("compare", {
  methods: ["POST"],
  authLevel: "function",
  route: "compare",
  handler: async (request, context) => {
    try {
      const body = await readJson(request);
      const result = await compareDocument(body, (m) => context.log(m));
      return json(200, result);
    } catch (e) {
      return fail(context, e);
    }
  },
});

app.http("publish", {
  methods: ["POST"],
  authLevel: "function",
  route: "publish",
  handler: async (request, context) => {
    try {
      const body = await readJson(request);
      const id = body?.standardId;
      if (!Number.isInteger(id) || id <= 0 || Object.keys(body).length !== 1) throw new UserError("Send { \"standardId\": <number> }.");
      const result = await publishStandard(id);
      if (!result) throw new UserError("No standard was found with that ID.", 404);
      return json(200, result);
    } catch (e) {
      return fail(context, e);
    }
  },
});

app.http("app", {
  methods: ["GET"],
  authLevel: "function",
  route: "app",
  handler: async () => {
    const nonce = randomBytes(16).toString("base64");
    const c = getConfig();
    const html = UI.replaceAll("__NONCE__", nonce)
      .replace("__CONTRACT_TYPES__", JSON.stringify(CONTRACT_TYPES))
      .replace("__CLAUSES_LIST_URL__", `https://${c.siteHost}${c.sitePath}/Lists/LDC_Clauses/AllItems.aspx`);
    return {
      status: 200,
      headers: {
        ...BASE_HEADERS,
        "Content-Type": "text/html; charset=utf-8",
        "Content-Security-Policy": `default-src 'none'; script-src 'nonce-${nonce}'; style-src 'nonce-${nonce}'; connect-src 'self'; img-src 'self' data:; base-uri 'none'; form-action 'none'; frame-ancestors 'none'`,
        "Permissions-Policy": "camera=(), microphone=(), geolocation=()",
      },
      body: html,
    };
  },
});
