// Checks AI answers and provides the keyword-based backup drafter.
const UNITS = { zero: 0, one: 1, two: 2, three: 3, four: 4, five: 5, six: 6, seven: 7, eight: 8, nine: 9, ten: 10, eleven: 11, twelve: 12, thirteen: 13, fourteen: 14, fifteen: 15, sixteen: 16, seventeen: 17, eighteen: 18, nineteen: 19 };
const TENS = { twenty: 20, thirty: 30, forty: 40, fifty: 50, sixty: 60, seventy: 70, eighty: 80, ninety: 90 };
const WORD_NUM = new RegExp(`\\b(?:(${Object.keys(TENS).join("|")})(?:[-\\s](${Object.keys(UNITS).join("|")}))?|(${Object.keys(UNITS).join("|")})|(hundred))\\b`, "gi");

// All numbers mentioned in a text, as digit strings ("forty-five" -> "45", "1,000" -> "1000").
export function numbersIn(text) {
  const out = new Set();
  for (const m of text.matchAll(/\d[\d,]*(?:\.\d+)?/g)) out.add(m[0].replace(/,/g, "").replace(/\.0+$/, ""));
  for (const m of text.matchAll(WORD_NUM)) {
    if (m[1]) out.add(String(TENS[m[1].toLowerCase()] + (m[2] ? UNITS[m[2].toLowerCase()] : 0)));
    else if (m[3]) out.add(String(UNITS[m[3].toLowerCase()]));
    else out.add("100");
  }
  return out;
}

export function checkAnswer(answer, clause, lists) {
  const problems = [];
  if (!lists.categories.includes(answer.category)) problems.push(`category '${answer.category}' is not on the list`);
  if (!lists.departments.includes(answer.department)) problems.push(`department '${answer.department}' is not on the list`);
  if (!lists.risks.includes(answer.risk)) problems.push(`risk '${answer.risk}' is not on the list`);
  if (!answer.title?.trim() || !answer.preferred?.trim()) problems.push("title or preferred position is empty");
  if (problems.length) return { valid: false, problems };

  const reasons = [];
  let confidence = Math.max(0, Math.min(1, Number(answer.confidence) || 0));
  const known = numbersIn(clause.text);
  for (const [field, label] of [["preferred", "Preferred"], ["fallback", "Fallback"], ["redLine", "Red line"]]) {
    const invented = [...numbersIn(answer[field] || "")].filter((n) => !known.has(n));
    if (invented.length) reasons.push(`${label} uses ${invented.map((n) => `'${n}'`).join(", ")}, which does not appear in the clause`);
  }
  if (reasons.length) confidence = Math.min(confidence, 0.45);
  if (confidence < 0.5 && !reasons.length) reasons.push("AI confidence below 0.5");
  if (answer.department === lists.fallbackDepartment) reasons.push(`No clear department, so it was set to ${lists.fallbackDepartment}`);

  const keywords = [...new Set((answer.keywords || []).map((k) => k.toLowerCase().trim()).filter(Boolean))].slice(0, 10);
  return {
    valid: true,
    draft: { ...answer, keywords, confidence: Math.round(confidence * 100) / 100, needsReview: reasons.length > 0, reviewReasons: reasons },
  };
}

// Backup method used when the AI answer is unusable. Deliberately simple and conservative.
const CATEGORY_HINTS = [
  ["Commercial", /\b(fee|fees|invoice|payment|price|pricing|charges?|late fee|tax)\b/i],
  ["Privacy", /\b(personal data|data protection|gdpr|privacy|data processing|processor)\b/i],
  ["Liability", /\b(liabilit|indemn|damages|limitation)\b/i],
  ["Term", /\b(term|terminat|renew|expir)\b/i],
  ["Confidentiality", /\b(confidential|non-disclosure|disclos)\b/i],
  ["Intellectual property", /\b(intellectual property|copyright|licen[cs]e|patent|trademark)\b/i],
  ["Warranty", /\b(warrant|represent)\b/i],
  ["Assignment", /\b(assign|subcontract|transfer)\b/i],
  ["Employment", /\b(employee|personnel|staff|non-solicit)\b/i],
  ["Delivery", /\b(deliver|milestone|acceptance|service level|sla)\b/i],
];
const DEPT_FOR = { Commercial: "Finance", Privacy: "Security / IT" };
const STOP = new Set("the a an and or of to in for on by with shall will may be is are any all this that such as from at its it party parties agreement".split(" "));

export function fallbackDraft(clause, lists, why) {
  const hit = CATEGORY_HINTS.find(([cat, re]) => lists.categories.includes(cat) && re.test(clause.text));
  const category = hit ? hit[0] : lists.categories.includes("General") ? "General" : lists.categories[0];
  const dept = DEPT_FOR[category] && lists.departments.includes(DEPT_FOR[category]) ? DEPT_FOR[category] : lists.departments.includes("Legal") ? "Legal" : lists.fallbackDepartment;
  const firstSentence = clause.text.split(/(?<=[.;])\s/)[0].slice(0, 400);
  const freq = new Map();
  for (const w of clause.text.toLowerCase().match(/[a-z][a-z-]{3,}/g) || []) if (!STOP.has(w)) freq.set(w, (freq.get(w) || 0) + 1);
  const keywords = [...freq.entries()].sort((a, b) => b[1] - a[1]).slice(0, 6).map(([w]) => w);
  return {
    title: clause.heading || category,
    category,
    department: dept,
    risk: lists.risks.includes("medium") ? "medium" : lists.risks[0],
    preferred: firstSentence,
    fallback: "Minor wording changes that keep the same rights and obligations, with Legal sign-off.",
    redLine: "Changes that shift risk to us, remove one of our rights or add an obligation without Legal approval.",
    keywords,
    confidence: 0.3,
    needsReview: true,
    reviewReasons: [`Drafted by the backup method because ${why}`],
  };
}
