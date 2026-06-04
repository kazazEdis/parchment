// Bookmarks (ARCHITECTURE.md §3.11). Parse w:bookmarkStart/End (range markers) for navigation /
// cross-references, and insert a bookmark over a selection. Word's hidden "_GoBack" is filtered out.
import { findElements, getAttr, escapeXmlAttr } from "./xml";
import { type Paragraph, type Run } from "./model";
import { formatRange } from "./edit";
import { emitRun, emitParagraphProps } from "./serialize";

export interface Bookmark {
  id: string;
  name: string;
}

/** All named bookmarks in the document, in order (excluding Word's _GoBack). */
export function parseBookmarks(documentXml: string): Bookmark[] {
  return findElements(documentXml, "w:bookmarkStart")
    .map((b) => ({ id: getAttr(b.openTag, "w:id") ?? "", name: getAttr(b.openTag, "w:name") ?? "" }))
    .filter((b) => b.name && b.name !== "_GoBack");
}

/** Insert a bookmark over a character range [start, end) of a paragraph. */
export function insertBookmark(documentXml: string, paragraph: Paragraph, start: number, end: number, id: string, name: string): string {
  if (start > end || !paragraph.children.every((n) => n.type === "run")) return documentXml;
  const startMark = `<w:bookmarkStart w:id="${escapeXmlAttr(id)}" w:name="${escapeXmlAttr(name)}"/>`;
  const endMark = `<w:bookmarkEnd w:id="${escapeXmlAttr(id)}"/>`;
  const runs = formatRange(paragraph, start, end, {}).children as Run[]; // split at boundaries
  const parts: string[] = [];
  let offset = 0;
  if (start === end) parts.push(startMark, endMark); // zero-length bookmark at the caret
  for (const r of runs) {
    if (start !== end && offset === start) parts.push(startMark);
    parts.push(emitRun(r));
    offset += r.text.length;
    if (start !== end && offset === end) parts.push(endMark);
  }
  const newOuter = `<w:p>${emitParagraphProps(paragraph.pPr)}${parts.join("")}</w:p>`;
  return documentXml.slice(0, paragraph.source.start) + newOuter + documentXml.slice(paragraph.source.end);
}
