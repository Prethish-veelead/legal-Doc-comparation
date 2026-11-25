// Creates the Part A SharePoint lists (safe to run more than once).
import { loadDotEnv } from "../src/lib/config.js";
loadDotEnv();
const { ensureLists } = await import("../src/lib/sharepoint.js");
await ensureLists();
console.log("Done.");
