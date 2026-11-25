// Groups text blocks into clauses. A clause starts at a heading or a numbered line.
const NUMBERED = /^((?:clause|section|article)\s+)?(\d+(?:\.\d+)*|[A-Z])[.)]?\s+(\S.*)$/i;
const SKIP_TITLE = /^(table of )?contents$|signature|in witness whereof|execution page|signed (by|for)/i;
const MIN_BODY = 40;
const MAX_CHARS = 6000;
const CHUNK_CHARS = 1500;

function startOf(block) {
  const m = block.text.match(NUMBERED);
  // A numbered line counts as a start; skip plain numbers inside sentences (e.g. "30 days").
  if (m && (m[1] || /\./.test(m[2]) || block.text.length < 120 || /^\d+[.)]\s/.test(block.text))) {
    const rest = m[3];
    const shortTitle = rest.length <= 90 && !/[.;:]$/.test(rest);
    return { ref: m[2], title: block.heading || shortTitle ? rest : null, inlineBody: block.heading || shortTitle ? "" : rest };
  }
  if (block.heading) return { ref: null, title: block.text, inlineBody: "" };
  return null;
}

export function splitClauses(blocks) {
  const sections = [];
  const skipped = [];
  let current = null;
  let preamble = [];

  for (const b of blocks) {
    const s = startOf(b);
    if (s) {
      current = { ref: s.ref, title: s.title, page: b.page, parts: s.inlineBody ? [s.inlineBody] : [] };
      sections.push(current);
    } else if (current) {
      current.parts.push(b.text);
    } else {
      preamble.push(b.text);
    }
  }

  // No structure at all (e.g. a plain README): chunk paragraphs.
  if (!sections.length) {
    let buf = [];
    const flush = () => {
      if (buf.length) sections.push({ ref: null, title: null, page: null, parts: buf });
      buf = [];
    };
    for (const t of preamble) {
      if (buf.join(" ").length + t.length > CHUNK_CHARS) flush();
      buf.push(t);
    }
    flush();
    preamble = [];
  }
  if (preamble.length) skipped.push("Cover page or preamble");

  const clauses = [];
  for (const s of sections) {
    const body = s.parts.join("\n").trim();
    const label = s.title || s.ref || "Untitled section";
    if (s.title && SKIP_TITLE.test(s.title)) {
      skipped.push(`${label} (not a clause)`);
      continue;
    }
    if (body.length < MIN_BODY) {
      // A heading followed directly by numbered sub-clauses has no body of its own.
      skipped.push(`${label} (heading only)`);
      continue;
    }
    clauses.push({
      sectionRef: [s.ref, s.page ? `page ${s.page}` : null].filter(Boolean).join(", ") || null,
      heading: s.title,
      text: body.length > MAX_CHARS ? body.slice(0, MAX_CHARS) : body,
    });
  }
  return { sectionsFound: sections.length, clauses, skipped };
}
