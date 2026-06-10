// Headless Document API — the high-level, whole-document surface (SuperDoc exposes `editor.doc`; this
// is the pure, Node-runnable, deterministic equivalent for document automation: open → query/edit →
// save, with every operation a plain function over the model + preserve-and-patch writer).
//
// This is where we beat SuperDoc for the offer/CPQ use case: template fills, batch redaction, and
// accept/reject run server-side with no browser, no ProseMirror, no AGPL — and `replaceText` is
// run-split aware (§edit.replaceInParagraph), which naive editors get wrong.
import { type DocxPackage, readDocx, writeDocx, docxToBlob, getPartText, setPartText } from "./opc";
import { parseDocument, documentText, type DocumentModel, type Block, type Paragraph, type Run, type Inline } from "./model";
import { parseStyles, type StyleSheet } from "./styles";
import { parseNumbering, type Numbering } from "./numbering";
import { patchAll, emitParagraph } from "./serialize";
import { replaceInParagraph, acceptChanges, rejectChanges } from "./edit";
import { nextCommentId, addComment, wrapParagraphComment, wrapCommentRange } from "./comments";

export interface Doc {
  pkg: DocxPackage;
  documentXml: string;
  /** Stable copy of comments.xml (so append-style edits stay pure under reducer double-invoke). */
  commentsXml: string;
  model: DocumentModel;
  styles: StyleSheet;
  numbering: Numbering;
}

/** Build a Doc from an already-read package. */
export function fromPackage(pkg: DocxPackage): Doc {
  const documentXml = getPartText(pkg, "word/document.xml") ?? "";
  return {
    pkg,
    documentXml,
    commentsXml: getPartText(pkg, "word/comments.xml") ?? "",
    model: parseDocument(documentXml),
    styles: parseStyles(getPartText(pkg, "word/styles.xml") ?? "", getPartText(pkg, "word/theme/theme1.xml")),
    numbering: parseNumbering(getPartText(pkg, "word/numbering.xml") ?? ""),
  };
}

/** Open a .docx (bytes/Blob/ArrayBuffer) into a Doc. */
export async function openDocx(input: ArrayBuffer | Uint8Array | Blob): Promise<Doc> {
  return fromPackage(await readDocx(input));
}

/** Every paragraph in reading order, descending into table cells. */
export function allParagraphs(model: DocumentModel): Paragraph[] {
  const acc: Paragraph[] = [];
  const walk = (blocks: Block[]): void => {
    for (const b of blocks) {
      if (b.type === "paragraph") acc.push(b);
      else for (const row of b.rows) for (const cell of row.cells) walk(cell.blocks);
    }
  };
  walk(model.body);
  return acc;
}

const hasTrackedRun = (p: Paragraph): boolean => {
  const any = (nodes: Inline[]): boolean =>
    nodes.some((n) => (n.type === "run" ? n.track !== undefined : n.type === "hyperlink" ? any(n.children) : false));
  return any(p.children);
};

/** Full document text (paragraphs joined by newlines, tables flattened). */
export function getText(doc: Doc): string {
  return documentText(doc.model);
}

/** Re-derive a Doc after writing new document.xml into the package (keeps the model in sync). */
function withDocumentXml(doc: Doc, newXml: string): Doc {
  setPartText(doc.pkg, "word/document.xml", newXml);
  return { ...doc, documentXml: newXml, model: parseDocument(newXml) };
}

/**
 * Replace every occurrence of `search` with `replace` across the whole document, run-split aware.
 * Returns the updated Doc and the number of replacements (0 ⇒ Doc unchanged).
 */
export function replaceText(doc: Doc, search: string, replace: string): { doc: Doc; count: number } {
  const edits: { span: { start: number; end: number }; xml: string }[] = [];
  let count = 0;
  for (const p of allParagraphs(doc.model)) {
    const r = replaceInParagraph(p, search, replace);
    if (r.count > 0) {
      edits.push({ span: p.source, xml: emitParagraph(r.paragraph) });
      count += r.count;
    }
  }
  if (count === 0) return { doc, count: 0 };
  return { doc: withDocumentXml(doc, patchAll(doc.documentXml, edits)), count };
}

/** Fill `{token}`-style placeholders from a map (run-split aware). Returns updated Doc + total fills. */
export function fillTemplate(doc: Doc, values: Record<string, string>, wrap: (k: string) => string = (k) => `{${k}}`): { doc: Doc; count: number } {
  let current = doc;
  let count = 0;
  for (const [key, value] of Object.entries(values)) {
    const r = replaceText(current, wrap(key), value);
    current = r.doc;
    count += r.count;
  }
  return { doc: current, count };
}

/** Apply a transform to every paragraph matching `predicate`, writing the results back. */
export function transformParagraphs(doc: Doc, predicate: (p: Paragraph) => boolean, transform: (p: Paragraph) => Paragraph): Doc {
  const edits = allParagraphs(doc.model)
    .filter(predicate)
    .map((p) => ({ span: p.source, xml: emitParagraph(transform(p)) }));
  return edits.length ? withDocumentXml(doc, patchAll(doc.documentXml, edits)) : doc;
}

/** Anchor a new comment on a paragraph: wrap its range in the body + append to comments.xml. */
export function addCommentToParagraph(doc: Doc, paragraphIndex: number, comment: { author?: string; initials?: string; text: string }): Doc {
  const p = allParagraphs(doc.model)[paragraphIndex];
  if (!p) return doc;
  // Read the stable doc.commentsXml (NOT the mutating pkg) so this stays pure under double-invoke.
  const id = nextCommentId(doc.commentsXml);
  const newComments = addComment(doc.commentsXml, { id, author: comment.author, initials: comment.initials, text: comment.text, date: new Date().toISOString() });
  const newXml = wrapParagraphComment(doc.documentXml, p, id);
  setPartText(doc.pkg, "word/comments.xml", newComments);
  setPartText(doc.pkg, "word/document.xml", newXml);
  return { ...doc, documentXml: newXml, commentsXml: newComments, model: parseDocument(newXml) };
}

/** Anchor a comment on a sub-paragraph character range [start, end). */
export function addCommentToRange(doc: Doc, paragraphIndex: number, start: number, end: number, comment: { author?: string; initials?: string; text: string }): Doc {
  const p = allParagraphs(doc.model)[paragraphIndex];
  if (!p) return doc;
  const id = nextCommentId(doc.commentsXml);
  const newComments = addComment(doc.commentsXml, { id, author: comment.author, initials: comment.initials, text: comment.text, date: new Date().toISOString() });
  const newXml = wrapCommentRange(doc.documentXml, p, start, end, id);
  setPartText(doc.pkg, "word/comments.xml", newComments);
  setPartText(doc.pkg, "word/document.xml", newXml);
  return { ...doc, documentXml: newXml, commentsXml: newComments, model: parseDocument(newXml) };
}

/** Accept every tracked change in the document. */
export function acceptAllChanges(doc: Doc): Doc {
  return transformParagraphs(doc, hasTrackedRun, acceptChanges);
}

/** Reject every tracked change in the document. */
export function rejectAllChanges(doc: Doc): Doc {
  return transformParagraphs(doc, hasTrackedRun, rejectChanges);
}

/** Serialize the document back to .docx bytes. */
export async function save(doc: Doc): Promise<Uint8Array> {
  return writeDocx(doc.pkg);
}

/** Serialize to a downloadable Blob (browser). */
export async function saveBlob(doc: Doc): Promise<Blob> {
  return docxToBlob(await save(doc));
}
