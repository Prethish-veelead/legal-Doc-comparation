// Turns a document buffer into an ordered list of text blocks. Everything stays in memory.
// Block: { text, heading: boolean, page: number|null }
import mammoth from "mammoth";
import { extractText, getDocumentProxy } from "unpdf";
import { UserError } from "./graph.js";

const TEXT_EXT = new Set(["txt", "md", "markdown"]);

export function detectType(fileName, buffer) {
  const ext = (fileName.split(".").pop() || "").toLowerCase();
  const head = buffer.subarray(0, 8);
  if (head.subarray(0, 4).toString("latin1") === "%PDF") return "pdf";
  if (ext === "doc" || (head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0)) {
    throw new UserError("Old Word .doc files are not supported. Open the file in Word, save it as .docx and try again.", 415);
  }
  if (head[0] === 0x50 && head[1] === 0x4b && ext === "docx") return "docx";
  if (TEXT_EXT.has(ext)) return ext === "txt" ? "txt" : "md";
  if (fileName.toLowerCase().startsWith("readme") && !buffer.includes(0)) return "md";
  throw new UserError("This file type is not supported. Use a Word (.docx), PDF, text or README file.", 415);
}

const decode = (s) =>
  s
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, " ")
    .trim();

async function fromDocx(buffer) {
  const html = (await mammoth.convertToHtml({ buffer })).value;
  const blocks = [];
  for (const m of html.matchAll(/<(h[1-6]|p|li)[^>]*>([\s\S]*?)<\/\1>/g)) {
    const text = decode(m[2]);
    if (text) blocks.push({ text, heading: m[1].startsWith("h"), page: null });
  }
  return blocks;
}

async function fromPdf(buffer) {
  const pdf = await getDocumentProxy(new Uint8Array(buffer));
  const { text: pages } = await extractText(pdf, { mergePages: false });
  if (!pages.join("").trim()) {
    throw new UserError("This PDF has no readable text (it may be a scan). Use a text-based PDF or a Word file.", 422);
  }
  // Lines repeated on most pages are headers/footers.
  const pageLines = pages.map((p) => p.split(/\r?\n/).map((l) => l.replace(/\s+/g, " ").trim()).filter(Boolean));
  const counts = new Map();
  pageLines.forEach((lines) => new Set(lines).forEach((l) => counts.set(l, (counts.get(l) || 0) + 1)));
  const repeated = (l) => pages.length >= 3 && counts.get(l) >= Math.ceil(pages.length * 0.6);
  const blocks = [];
  pageLines.forEach((lines, i) => {
    for (const l of lines) {
      if (repeated(l) || /^(page\s*)?\d+(\s*(of|\/)\s*\d+)?$/i.test(l)) continue;
      blocks.push({ text: l, heading: false, page: i + 1 });
    }
  });
  return blocks;
}

function fromText(buffer, markdown) {
  const blocks = [];
  let para = [];
  const flush = () => {
    if (para.length) blocks.push({ text: para.join(" "), heading: false, page: null });
    para = [];
  };
  for (const raw of buffer.toString("utf8").split(/\r?\n/)) {
    const line = raw.trim();
    const h = markdown && line.match(/^#{1,6}\s+(.*)$/);
    if (h) {
      flush();
      blocks.push({ text: h[1].replace(/[*_`]/g, "").trim(), heading: true, page: null });
    } else if (!line || /^(```|---|\*\*\*)/.test(line)) {
      flush();
    } else {
      para.push(markdown ? line.replace(/^[-*+]\s+/, "").replace(/[*_`]/g, "") : line);
    }
  }
  flush();
  return blocks;
}

export async function extractBlocks(fileName, buffer) {
  const type = detectType(fileName, buffer);
  let blocks;
  if (type === "docx") blocks = await fromDocx(buffer);
  else if (type === "pdf") blocks = await fromPdf(buffer);
  else blocks = fromText(buffer, type === "md");
  if (!blocks.length) throw new UserError("No text was found in this document.", 422);
  return { type, blocks, pages: type === "pdf" ? Math.max(...blocks.map((b) => b.page || 0)) : null };
}
