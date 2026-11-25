// Runs Part B once from the command line, without the Functions host. Saves nothing to SharePoint.
// Usage: npm run run-compare -- "<their document link>" <standardId> "<our role>"
//    or: npm run run-compare -- --file <local test file> <standardId> "<our role>"
import { readFileSync } from "node:fs";
import { basename } from "node:path";
import { loadDotEnv } from "../src/lib/config.js";
loadDotEnv();
const { compareDocument } = await import("../src/lib/compare.js");

const args = process.argv.slice(2);
const fileMode = args[0] === "--file";
if (fileMode) args.shift();
const [source, standardId, ourRole = ""] = args;
const documentUrl = fileMode ? `local:${basename(source || "")}` : source;
if (!source || !standardId) {
  console.error('Usage: npm run run-compare -- "<their document link>" <standardId> "<our role>"  (or --file <path> instead of the link)');
  process.exit(1);
}
const localDoc = fileMode ? { fileName: basename(source), buffer: readFileSync(source) } : null;
try {
  const result = await compareDocument({ documentUrl, standardId: Number(standardId), ourRole }, (m) => console.error("·", m), localDoc);
  console.log(JSON.stringify(result, null, 2));
} catch (e) {
  console.error("Failed:", e.message);
  process.exit(1);
}
