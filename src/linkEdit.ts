// Apply a hyperlink to a sub-paragraph selection. Splits the runs at the range boundaries and wraps
// them in w:hyperlink (the external target is a relationship; the caller registers it via
// opcParts.addRelationship and passes the rId, so this stays a pure string transform).
import { type Paragraph, type Run } from "./model";
import { formatRange } from "./edit";
import { emitRun, emitParagraphProps } from "./serialize";
import { escapeXmlAttr } from "./xml";

export function wrapHyperlink(documentXml: string, paragraph: Paragraph, start: number, end: number, rId: string): string {
  if (start >= end || !paragraph.children.every((n) => n.type === "run")) return documentXml;
  const runs = formatRange(paragraph, start, end, {}).children as Run[]; // split at start + end
  const parts: string[] = [];
  let offset = 0;
  let open = false;
  for (const r of runs) {
    if (offset === start) { parts.push(`<w:hyperlink r:id="${escapeXmlAttr(rId)}">`); open = true; }
    parts.push(emitRun(r));
    offset += r.text.length;
    if (offset === end && open) { parts.push(`</w:hyperlink>`); open = false; }
  }
  if (open) parts.push(`</w:hyperlink>`);
  const newOuter = `<w:p>${emitParagraphProps(paragraph.pPr)}${parts.join("")}</w:p>`;
  return documentXml.slice(0, paragraph.source.start) + newOuter + documentXml.slice(paragraph.source.end);
}
