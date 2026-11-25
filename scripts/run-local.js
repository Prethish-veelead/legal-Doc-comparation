// Runs Part A once from the command line, without the Functions host.
// Usage: npm run run-local -- "<document link>" "<contract type>" "<our role>" [--force]
import { loadDotEnv } from "../src/lib/config.js";
loadDotEnv();
const { processDocument } = await import("../src/lib/pipeline.js");

const args = process.argv.slice(2);
const force = args.includes("--force");
const [documentUrl, contractType = "Master services", ourRole = ""] = args.filter((a) => a !== "--force");
if (!documentUrl) {
  console.error('Usage: npm run run-local -- "<document link>" "<contract type>" "<our role>" [--force]');
  process.exit(1);
}
try {
  const result = await processDocument({ documentUrl, contractType, ourRole, force }, (m) => console.error("·", m));
  console.log(JSON.stringify(result, null, 2));
} catch (e) {
  console.error("Failed:", e.message);
  process.exit(1);
}
