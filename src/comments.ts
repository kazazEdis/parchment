// Comments (ARCHITECTURE.md §3.13). The body marks a commented span with w:commentRangeStart/End
// (matched by w:id) and a w:commentReference; the comment bodies live in word/comments.xml. This
// module parses the comment definitions and locates the commented text in the body — enough to
// render a comments panel + highlight the ranges. (Threading/people live in commentsExtended.xml /
// people.xml; authoring + the 4-file write linkage is the next step.)
import { findElement, findElements, getAttr, unescapeXml, escapeXmlAttr, escapeXmlText } from "./xml";
import type { Paragraph, Run } from "./model";
import { formatRange } from "./edit";
import { emitRun, emitParagraphProps } from "./serialize";

export interface CommentDef {
  id: string;
  author?: string;
  initials?: string;
  date?: string;
  /** The comment body text (paragraphs flattened). */
  text: string;
}

const textOf = (xml: string, from: number, to: number): string =>
  findElements(xml.slice(from, to), "w:t").map((t) => unescapeXml(xml.slice(from, to).slice(t.innerStart, t.innerEnd))).join("");

/** Parse word/comments.xml into comment definitions. */
export function parseComments(commentsXml: string | undefined): CommentDef[] {
  if (!commentsXml) return [];
  return findElements(commentsXml, "w:comment").map((c) => ({
    id: getAttr(c.openTag, "w:id") ?? "",
    author: getAttr(c.openTag, "w:author"),
    initials: getAttr(c.openTag, "w:initials"),
    date: getAttr(c.openTag, "w:date"),
    text: textOf(commentsXml, c.innerStart, c.innerEnd),
  }));
}

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** Next free numeric comment id for a comments.xml part. */
export function nextCommentId(commentsXml: string | undefined): string {
  let max = -1;
  if (commentsXml) {
    for (const c of findElements(commentsXml, "w:comment")) {
      const id = parseInt(getAttr(c.openTag, "w:id") ?? "", 10);
      if (Number.isFinite(id)) max = Math.max(max, id);
    }
  }
  return String(max + 1);
}

/** Append a comment definition to comments.xml (creating the part if absent). */
export function addComment(commentsXml: string | undefined, def: CommentDef): string {
  const attrs =
    `w:id="${escapeXmlAttr(def.id)}"` +
    (def.author !== undefined ? ` w:author="${escapeXmlAttr(def.author)}"` : "") +
    (def.initials !== undefined ? ` w:initials="${escapeXmlAttr(def.initials)}"` : "") +
    (def.date !== undefined ? ` w:date="${escapeXmlAttr(def.date)}"` : "");
  const commentXml = `<w:comment ${attrs}><w:p><w:r><w:t xml:space="preserve">${escapeXmlText(def.text)}</w:t></w:r></w:p></w:comment>`;
  const root = commentsXml ? findElement(commentsXml, "w:comments") : undefined;
  if (!commentsXml || !root) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ${W_NS}>${commentXml}</w:comments>`;
  }
  return commentsXml.slice(0, root.innerEnd) + commentXml + commentsXml.slice(root.innerEnd);
}

/**
 * Wrap a whole paragraph's content in document.xml with a comment range (commentRangeStart/End +
 * commentReference) for comment `id`. Returns the new document.xml. (Whole-paragraph anchoring in v1;
 * sub-paragraph ranges would split runs at the offsets, like edit.spliceRunRange.)
 */
export function wrapParagraphComment(documentXml: string, paragraph: Paragraph, id: string): string {
  const outer = documentXml.slice(paragraph.source.start, paragraph.source.end);
  const pEl = findElement(outer, "w:p");
  if (!pEl) return documentXml;
  const inner = outer.slice(pEl.innerStart, pEl.innerEnd);
  const pPr = findElement(inner, "w:pPr");
  const at = pPr ? pPr.outerEnd : 0; // after pPr (or at content start)
  const startMark = `<w:commentRangeStart w:id="${escapeXmlAttr(id)}"/>`;
  const endMark = `<w:commentRangeEnd w:id="${escapeXmlAttr(id)}"/><w:r><w:commentReference w:id="${escapeXmlAttr(id)}"/></w:r>`;
  const newInner = inner.slice(0, at) + startMark + inner.slice(at) + endMark;
  const newOuter = outer.slice(0, pEl.innerStart) + newInner + outer.slice(pEl.innerEnd);
  return documentXml.slice(0, paragraph.source.start) + newOuter + documentXml.slice(paragraph.source.end);
}

/**
 * Anchor a comment on a sub-paragraph character range [start, end). Splits the runs at the range
 * boundaries (via formatRange) and inserts commentRangeStart/End + reference at those boundaries.
 * Falls back to whole-paragraph wrapping for non-run-only paragraphs or a full/empty range.
 */
export function wrapCommentRange(documentXml: string, paragraph: Paragraph, start: number, end: number, id: string): string {
  if (start >= end || !paragraph.children.every((n) => n.type === "run")) return wrapParagraphComment(documentXml, paragraph, id);
  const runs = (formatRange(paragraph, start, end, {}).children as Run[]); // split at start + end
  const startMark = `<w:commentRangeStart w:id="${escapeXmlAttr(id)}"/>`;
  const endMark = `<w:commentRangeEnd w:id="${escapeXmlAttr(id)}"/><w:r><w:commentReference w:id="${escapeXmlAttr(id)}"/></w:r>`;
  const parts: string[] = [];
  let offset = 0;
  for (const r of runs) {
    if (offset === start) parts.push(startMark);
    parts.push(emitRun(r));
    offset += r.text.length;
    if (offset === end) parts.push(endMark);
  }
  const newOuter = `<w:p>${emitParagraphProps(paragraph.pPr)}${parts.join("")}</w:p>`;
  return documentXml.slice(0, paragraph.source.start) + newOuter + documentXml.slice(paragraph.source.end);
}

/** For each comment id, the body text it annotates (between w:commentRangeStart/End). */
export function commentRanges(documentXml: string): Map<string, string> {
  const out = new Map<string, string>();
  for (const start of findElements(documentXml, "w:commentRangeStart")) {
    const id = getAttr(start.openTag, "w:id");
    if (id == null) continue;
    let end: { outerStart: number; outerEnd: number } | undefined;
    let cursor = start.outerEnd;
    for (;;) {
      const e = findElement(documentXml, "w:commentRangeEnd", { from: cursor });
      if (!e) break;
      if (getAttr(e.openTag, "w:id") === id) { end = e; break; }
      cursor = e.outerEnd;
    }
    if (!end) continue;
    out.set(id, textOf(documentXml, start.outerEnd, end.outerStart));
  }
  return out;
}
