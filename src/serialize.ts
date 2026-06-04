// Preserve-and-patch writer (ARCHITECTURE.md §3.9). The fidelity contract on write: rewrite ONLY
// the nodes the user edited, splicing fresh XML into the original document.xml at each node's
// recorded source span, so every untouched byte survives. We never re-serialise the whole part.
//
// The model is intentionally sparse, so re-emitting a node is lossy for content the model does not
// carry (fields, symbols, drawings). That is acceptable precisely because we only re-emit EDITED
// nodes — an untouched paragraph keeps its original bytes verbatim. Editing a paragraph that
// contains a field/drawing in v1 would drop it; the editor should gate such paragraphs (or carry
// their raw inline XML) until those inlines are modelled. Plain text/run edits round-trip cleanly.
import { escapeXmlText, escapeXmlAttr } from "./xml";
import { pointsToHalfPoints } from "./units";
import type { RunProps, ParagraphProps } from "./props";
import type { Paragraph, Run, Inline, SourceSpan } from "./model";

// ── run properties → <w:rPr> ────────────────────────────────────────────────────────────────────
const toggleEl = (tag: string, on: boolean): string => (on ? `<${tag}/>` : `<${tag} w:val="false"/>`);

export function emitRunProps(p: RunProps): string {
  const parts: string[] = [];
  if (p.styleId !== undefined) parts.push(`<w:rStyle w:val="${escapeXmlAttr(p.styleId)}"/>`);
  if (p.fonts) {
    const a: string[] = [];
    if (p.fonts.ascii) a.push(`w:ascii="${escapeXmlAttr(p.fonts.ascii)}"`);
    if (p.fonts.hAnsi) a.push(`w:hAnsi="${escapeXmlAttr(p.fonts.hAnsi)}"`);
    if (p.fonts.eastAsia) a.push(`w:eastAsia="${escapeXmlAttr(p.fonts.eastAsia)}"`);
    if (p.fonts.cs) a.push(`w:cs="${escapeXmlAttr(p.fonts.cs)}"`);
    if (a.length) parts.push(`<w:rFonts ${a.join(" ")}/>`);
  }
  if (p.bold !== undefined) parts.push(toggleEl("w:b", p.bold));
  if (p.italic !== undefined) parts.push(toggleEl("w:i", p.italic));
  if (p.caps !== undefined) parts.push(toggleEl("w:caps", p.caps));
  if (p.smallCaps !== undefined) parts.push(toggleEl("w:smallCaps", p.smallCaps));
  if (p.strike !== undefined) parts.push(toggleEl("w:strike", p.strike));
  if (p.color !== undefined) parts.push(`<w:color w:val="${escapeXmlAttr(p.color)}"/>`);
  if (p.fontSize !== undefined) parts.push(`<w:sz w:val="${pointsToHalfPoints(p.fontSize)}"/>`);
  if (p.highlight !== undefined) parts.push(`<w:highlight w:val="${escapeXmlAttr(p.highlight)}"/>`);
  if (p.underline !== undefined) parts.push(`<w:u w:val="${escapeXmlAttr(p.underline)}"/>`);
  if (p.vertAlign !== undefined) parts.push(`<w:vertAlign w:val="${p.vertAlign}"/>`);
  if (p.shading !== undefined) parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${escapeXmlAttr(p.shading)}"/>`);
  return parts.length ? `<w:rPr>${parts.join("")}</w:rPr>` : "";
}

// ── paragraph properties → <w:pPr> ──────────────────────────────────────────────────────────────
export function emitParagraphProps(p: ParagraphProps): string {
  const parts: string[] = [];
  if (p.styleId !== undefined) parts.push(`<w:pStyle w:val="${escapeXmlAttr(p.styleId)}"/>`);
  if (p.keepNext) parts.push(`<w:keepNext/>`);
  if (p.keepLines) parts.push(`<w:keepLines/>`);
  if (p.numbering) parts.push(`<w:numPr><w:ilvl w:val="${p.numbering.level}"/><w:numId w:val="${p.numbering.numId}"/></w:numPr>`);
  if (p.shading !== undefined) parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${escapeXmlAttr(p.shading)}"/>`);
  if (p.spacing) {
    const a: string[] = [];
    if (p.spacing.before !== undefined) a.push(`w:before="${p.spacing.before}"`);
    if (p.spacing.after !== undefined) a.push(`w:after="${p.spacing.after}"`);
    if (p.spacing.line !== undefined) a.push(`w:line="${p.spacing.line}"`);
    if (p.spacing.lineRule !== undefined) a.push(`w:lineRule="${escapeXmlAttr(p.spacing.lineRule)}"`);
    if (a.length) parts.push(`<w:spacing ${a.join(" ")}/>`);
  }
  if (p.indent) {
    const a: string[] = [];
    if (p.indent.left !== undefined) a.push(`w:left="${p.indent.left}"`);
    if (p.indent.right !== undefined) a.push(`w:right="${p.indent.right}"`);
    if (p.indent.firstLine !== undefined) a.push(`w:firstLine="${p.indent.firstLine}"`);
    if (p.indent.hanging !== undefined) a.push(`w:hanging="${p.indent.hanging}"`);
    if (a.length) parts.push(`<w:ind ${a.join(" ")}/>`);
  }
  if (p.alignment !== undefined) parts.push(`<w:jc w:val="${p.alignment}"/>`);
  if (p.outlineLevel !== undefined) parts.push(`<w:outlineLvl w:val="${p.outlineLevel}"/>`);
  if (p.markRunProps) {
    const r = emitRunProps(p.markRunProps);
    if (r) parts.push(r);
  }
  return parts.length ? `<w:pPr>${parts.join("")}</w:pPr>` : "";
}

// ── runs / inlines → XML ────────────────────────────────────────────────────────────────────────
function emitRunContent(text: string, isDel = false): string {
  const tag = isDel ? "w:delText" : "w:t"; // deleted runs store text in w:delText
  let out = "";
  let buf = "";
  const flush = (): void => {
    if (buf) { out += `<${tag} xml:space="preserve">${escapeXmlText(buf)}</${tag}>`; buf = ""; }
  };
  for (const ch of text) {
    if (ch === "\t") { flush(); out += "<w:tab/>"; }
    else if (ch === "\n") { flush(); out += "<w:br/>"; }
    else buf += ch;
  }
  flush();
  return out || `<${tag} xml:space="preserve"></${tag}>`;
}

export function emitRun(r: Run): string {
  const inner = `<w:r>${emitRunProps(r.rPr)}${emitRunContent(r.text, r.track?.type === "del")}</w:r>`;
  if (!r.track) return inner;
  const t = r.track;
  const tag = t.type === "ins" ? "w:ins" : "w:del";
  const attrs =
    ` w:id="${t.id ?? "0"}"` +
    (t.author !== undefined ? ` w:author="${escapeXmlAttr(t.author)}"` : "") +
    (t.date !== undefined ? ` w:date="${escapeXmlAttr(t.date)}"` : "");
  return `<${tag}${attrs}>${inner}</${tag}>`;
}

export function emitInline(n: Inline): string {
  if (n.type === "run") return emitRun(n);
  if (n.type === "hyperlink") {
    const attrs =
      (n.rId !== undefined ? ` r:id="${escapeXmlAttr(n.rId)}"` : "") +
      (n.anchor !== undefined ? ` w:anchor="${escapeXmlAttr(n.anchor)}"` : "");
    return `<w:hyperlink${attrs}>${n.children.map(emitInline).join("")}</w:hyperlink>`;
  }
  if (n.type === "footnoteRef") return `<w:r><w:footnoteReference w:id="${escapeXmlAttr(n.id)}"/></w:r>`;
  if (n.type === "math") return n.omml; // kept verbatim
  // drawing — not re-emittable from the sparse v1 model (see file header). Dropped on re-emit.
  return "";
}

export function emitParagraph(p: Paragraph): string {
  return `<w:p>${emitParagraphProps(p.pPr)}${p.children.map(emitInline).join("")}</w:p>`;
}

// ── splice (preserve-and-patch) ─────────────────────────────────────────────────────────────────

/** Replace the XML at `span` with `replacement`, leaving every other byte untouched. */
export function patchSpan(documentXml: string, span: SourceSpan, replacement: string): string {
  return documentXml.slice(0, span.start) + replacement + documentXml.slice(span.end);
}

/** Re-emit `paragraph` and splice it back at its recorded source span. */
export function patchParagraph(documentXml: string, paragraph: Paragraph): string {
  return patchSpan(documentXml, paragraph.source, emitParagraph(paragraph));
}

/**
 * Apply many block edits in one pass. Edits are spliced right-to-left (descending source offset) so
 * earlier offsets stay valid as later ones are replaced. Each edit pairs a source span with its new
 * XML (typically `emitParagraph(editedParagraph)`).
 */
export function patchAll(documentXml: string, edits: { span: SourceSpan; xml: string }[]): string {
  let out = documentXml;
  for (const e of [...edits].sort((a, b) => b.span.start - a.span.start)) {
    out = patchSpan(out, e.span, e.xml);
  }
  return out;
}

/** Pure model transform: a copy of `p` whose text is one run carrying the first run's formatting. */
export function setParagraphText(p: Paragraph, text: string): Paragraph {
  const firstRun = p.children.find((n): n is Run => n.type === "run");
  return { ...p, children: [{ type: "run", rPr: firstRun ? firstRun.rPr : {}, text }] };
}
