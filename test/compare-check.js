// Live check of Part B against drafts with known answers (calls Azure OpenAI, reads SharePoint, writes nothing).
// Usage: node test/compare-check.js [runs per draft, default 1]
import { readFileSync } from "node:fs";
import { loadDotEnv } from "../src/lib/config.js";
loadDotEnv();
const { compareDocument } = await import("../src/lib/compare.js");

const STANDARD_ID = 1; // Vendor MSA Standards v4
const CASES = {
  "counterparty-msa-bad.md": { "LIAB-01": "crosses_red_line", "PAY-02": "within_fallback", "IND-01": "crosses_red_line", "PRIV-01": "crosses_red_line", "TERM-02": "within_fallback", "REN-01": "crosses_red_line", "LAW-01": "missing" },
  // Exact limits: 30 days = fallback, 1.5% is not "above 1.5%", 75 days is past the 60-day fallback but not over 90.
  "counterparty-msa-borderline.md": { "LIAB-01": "outside_fallback", "PAY-02": ["within_fallback", "outside_fallback"], "IND-01": "missing", "PRIV-01": "meets_preferred", "TERM-02": "outside_fallback", "REN-01": "within_fallback", "LAW-01": "within_fallback" },
  "counterparty-msa-good.md": { "LIAB-01": "meets_preferred", "PAY-02": "meets_preferred", "IND-01": "meets_preferred", "PRIV-01": "meets_preferred", "TERM-02": "meets_preferred", "REN-01": "meets_preferred", "LAW-01": "meets_preferred" },
};

const runs = Number(process.argv[2] || 1);
let pass = 0;
let total = 0;
for (const [file, expected] of Object.entries(CASES)) {
  const buffer = readFileSync(new URL(`./fixtures/${file}`, import.meta.url));
  for (let i = 1; i <= runs; i++) {
    const r = await compareDocument({ documentUrl: `local:${file}`, standardId: STANDARD_ID, ourRole: "Customer" }, () => {}, { fileName: file, buffer });
    for (const c of r.checks) {
      total++;
      if ([expected[c.ruleCode]].flat().includes(c.verdict)) pass++;
      else {
        const f = r.findingsJson.findings.find((x) => x.ruleCode === c.ruleCode);
        console.log(`MISS ${file} run ${i} ${c.ruleCode}: got ${c.verdict}, expected ${expected[c.ruleCode]}; flagged for review: ${f?.needsReview ? "yes" : "NO"}`);
        if (!f?.needsReview) console.log(`     AI reasoning: ${f?.aiReasoning}`);
      }
    }
    console.log(`${file} run ${i}: ${r.run.durationSec}s, $${r.run.estimatedCostUsd}`);
  }
}
console.log(`\n${pass}/${total} verdicts correct`);
process.exit(pass === total ? 0 : 1);
