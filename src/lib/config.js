// Reads settings from environment variables (Function App settings in Azure,
// local.settings.json under `func start`, or .env for the scripts).
const REQUIRED = [
  "AZURE_TENANT_ID",
  "AZURE_CLIENT_ID",
  "AZURE_CLIENT_SECRET",
  "SHAREPOINT_SITE_URL",
  "SHAREPOINT_LIBRARY",
  "AZURE_OPENAI_ENDPOINT",
  "AZURE_OPENAI_API_KEY",
  "AZURE_OPENAI_DEPLOYMENT",
  "AZURE_OPENAI_API_VERSION",
];

let cached;

export function getConfig() {
  if (cached) return cached;
  const missing = REQUIRED.filter((k) => !process.env[k]);
  if (missing.length) throw new Error(`Missing required settings: ${missing.join(", ")}`);
  const site = new URL(process.env.SHAREPOINT_SITE_URL);
  cached = {
    tenantId: process.env.AZURE_TENANT_ID,
    clientId: process.env.AZURE_CLIENT_ID,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
    siteHost: site.hostname.toLowerCase(),
    sitePath: site.pathname.replace(/\/+$/, ""), // e.g. /sites/LegalDocComparison
    library: process.env.SHAREPOINT_LIBRARY,
    aoaiEndpoint: process.env.AZURE_OPENAI_ENDPOINT.replace(/\/+$/, ""),
    aoaiKey: process.env.AZURE_OPENAI_API_KEY,
    aoaiDeployment: process.env.AZURE_OPENAI_DEPLOYMENT,
    // Part B (comparison) model; falls back to the Part A model.
    compareDeployment: process.env.AZURE_OPENAI_COMPARE_DEPLOYMENT || process.env.AZURE_OPENAI_DEPLOYMENT,
    aoaiApiVersion: process.env.AZURE_OPENAI_API_VERSION,
    maxFileBytes: Number(process.env.MAX_FILE_MB || 20) * 1024 * 1024,
    aiConcurrency: Number(process.env.AI_CONCURRENCY || 3),
  };
  return cached;
}

// For scripts run with plain `node`: load .env from the project root if present.
export function loadDotEnv() {
  try {
    process.loadEnvFile(new URL("../../.env", import.meta.url));
  } catch (e) {
    if (e.code !== "ENOENT") throw e;
  }
}
