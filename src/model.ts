// The document model — document.xml parsed into an editable tree (ARCHITECTURE.md §3.3).
//
// Per "store raw, resolve late": nodes carry their *direct* (unresolved) pPr/rPr exactly as parsed;
// the render layer composes docDefaults → paragraph style → character style → these direct props.
// Every block records its absolute `source` span (offsets into document.xml) so the serializer can
// rewrite only edited nodes and leave the rest byte-identical (preserve-and-patch, §3.9).
//
// Read-only for now: tracked insertions (w:ins) are unwrapped (accepted), tracked deletions (w:del)
// and range markers (bookmarks, comments) are dropped from the read model — they become first-class
// when the track-changes / comments features land.
import { childElements, findElement, getAttr, localName, unescapeXml, type ElementSpan } from "./xml";
import { type ParagraphProps, type RunProps, parseParagraphProps, parseRunProps } from "./props";
import { parseMeasure } from "./units";

export interface SourceSpan {
  start: number;
  end: number;
}

/** A tracked-change wrapper (w:ins / w:del) carrying its revision metadata. */
export interface TrackChange {
  type: "ins" | "del";
  id?: string;
  author?: string;
  date?: string;
}

export interface Run {
  type: "run";
  rPr: RunProps;
  /** Run text; tabs become "\t", line/page breaks and carriage returns become "\n". */
  text: string;
  /** Present when this run sits inside a tracked insertion/deletion (w:ins/w:del). */
  track?: TrackChange;
}

export interface Drawing {
  type: "drawing";
  anchored: boolean;
  /** Relationship id of the image blip (r:embed), resolved to bytes via document.xml.rels. */
  rEmbed?: string;
  widthEmu?: number;
  heightEmu?: number;
  alt?: string;
}

export interface Hyperlink {
  type: "hyperlink";
  /** External target via rels (r:id), or… */
  rId?: string;
  /** …internal bookmark anchor (w:anchor). */
  anchor?: string;
  children: Inline[];
}

/** A w:footnoteReference marker; the footnote body lives in word/footnotes.xml. */
export interface FootnoteRef {
  type: "footnoteRef";
  id: string;
}

/** An Office Math (OMML) region; the original m:oMath XML is kept verbatim (render-only). */
export interface MathInline {
  type: "math";
  omml: string;
}

export type Inline = Run | Drawing | Hyperlink | FootnoteRef | MathInline;

export interface Paragraph {
  type: "paragraph";
  pPr: ParagraphProps;
  children: Inline[];
  source: SourceSpan;
}

export interface TableCellProps {
  gridSpan?: number;
  vMerge?: "restart" | "continue";
  width?: { value: number; type: string };
}

export interface TableCell {
  props: TableCellProps;
  blocks: Block[];
  source: SourceSpan;
}

export interface TableRow {
  cells: TableCell[];
  isHeader: boolean;
  source: SourceSpan;
}

export interface Table {
  type: "table";
  /** Column widths in twips, from w:tblGrid. */
  grid: number[];
  rows: TableRow[];
  source: SourceSpan;
}

export type Block = Paragraph | Table;

export interface SectionProps {
  pageSize?: { width: number; height: number; orient?: string };
  margins?: { top?: number; right?: number; bottom?: number; left?: number; header?: number; footer?: number; gutter?: number };
  cols?: { num: number; space?: number };
  headerRefs?: { type: string; rId: string }[];
  footerRefs?: { type: string; rId: string }[];
  titlePage?: boolean;
}

export interface DocumentModel {
  body: Block[];
  section?: SectionProps;
}

const inner = (xml: string, el: ElementSpan): string => xml.slice(el.innerStart, el.innerEnd);

/** Parse the body of document.xml into the document model. */
export function parseDocument(documentXml: string): DocumentModel {
  const bodyEl = findElement(documentXml, "w:body");
  if (!bodyEl) return { body: [] };
  const body = parseBlocks(documentXml, bodyEl.innerStart, bodyEl.innerEnd);
  const sectEl = findElement(documentXml, "w:sectPr", { from: bodyEl.innerStart, to: bodyEl.innerEnd });
  return { body, section: sectEl ? parseSectPr(documentXml, sectEl) : undefined };
}

/** Parse the blocks of any container root (w:hdr / w:ftr / w:body) — used for headers/footers. */
export function parseContainer(xml: string, rootName: string): Block[] {
  const el = findElement(xml, rootName);
  return el ? parseBlocks(xml, el.innerStart, el.innerEnd) : [];
}

function parseBlocks(xml: string, from: number, to: number): Block[] {
  const out: Block[] = [];
  for (const el of childElements(xml, from, to)) {
    const ln = localName(el.name);
    if (ln === "p") out.push(parseParagraph(xml, el));
    else if (ln === "tbl") out.push(parseTable(xml, el));
    // sectPr (handled separately), bookmarkStart/End, sdt at block level → skipped in v1
  }
  return out;
}

function parseParagraph(xml: string, el: ElementSpan): Paragraph {
  const pPrEl = findElement(xml, "w:pPr", { from: el.innerStart, to: el.innerEnd });
  const pPr = pPrEl ? parseParagraphProps(inner(xml, pPrEl)) : {};
  return {
    type: "paragraph",
    pPr,
    children: parseInlines(xml, el.innerStart, el.innerEnd),
    source: { start: el.outerStart, end: el.outerEnd },
  };
}

function trackOf(el: ElementSpan, type: "ins" | "del"): TrackChange {
  return { type, id: getAttr(el.openTag, "w:id"), author: getAttr(el.openTag, "w:author"), date: getAttr(el.openTag, "w:date") };
}

function parseInlines(xml: string, from: number, to: number, track?: TrackChange): Inline[] {
  const out: Inline[] = [];
  for (const el of childElements(xml, from, to)) {
    const ln = localName(el.name);
    if (ln === "r") parseRunInto(xml, el, out, track);
    else if (ln === "hyperlink") {
      out.push({
        type: "hyperlink",
        rId: getAttr(el.openTag, "r:id"),
        anchor: getAttr(el.openTag, "w:anchor"),
        children: parseInlines(xml, el.innerStart, el.innerEnd, track),
      });
    } else if (ln === "ins" || ln === "del") {
      // Tracked change: tag the contained runs with the revision metadata.
      for (const c of parseInlines(xml, el.innerStart, el.innerEnd, trackOf(el, ln))) out.push(c);
    } else if (ln === "smartTag" || ln === "sdt" || ln === "fldSimple") {
      // Unwrap wrappers + simple fields → their cached result runs render as content.
      for (const c of parseInlines(xml, el.innerStart, el.innerEnd, track)) out.push(c);
    } else if (ln === "oMath") {
      out.push({ type: "math", omml: xml.slice(el.outerStart, el.outerEnd) });
    }
    // pPr, bookmark*, commentRange*, proofErr → skipped in the v1 read model
  }
  return out;
}

function parseRunInto(xml: string, el: ElementSpan, out: Inline[], track?: TrackChange): void {
  const rPrEl = findElement(xml, "w:rPr", { from: el.innerStart, to: el.innerEnd });
  const rPr = rPrEl ? parseRunProps(inner(xml, rPrEl)) : {};
  let buf = "";
  const flush = (): void => {
    if (buf) {
      const run: Run = { type: "run", rPr, text: buf };
      if (track) run.track = track;
      out.push(run);
      buf = "";
    }
  };
  for (const c of childElements(xml, el.innerStart, el.innerEnd)) {
    const ln = localName(c.name);
    if (ln === "t" || ln === "delText") buf += unescapeXml(inner(xml, c)); // delText: deleted run's text
    else if (ln === "tab") buf += "\t";
    else if (ln === "cr" || ln === "br") buf += "\n";
    else if (ln === "drawing") { flush(); out.push(parseDrawing(xml, c)); }
    else if (ln === "footnoteReference") { flush(); out.push({ type: "footnoteRef", id: getAttr(c.openTag, "w:id") ?? "" }); }
    // rPr, sym, noBreakHyphen, fldChar/instrText → skipped in v1
  }
  flush();
}

function parseDrawing(xml: string, el: ElementSpan): Drawing {
  const body = inner(xml, el);
  const extent = findElement(body, "wp:extent");
  const blip = findElement(body, "a:blip");
  const docPr = findElement(body, "wp:docPr");
  return {
    type: "drawing",
    anchored: findElement(body, "wp:anchor") !== undefined,
    rEmbed: blip ? getAttr(blip.openTag, "r:embed") : undefined,
    widthEmu: extent ? parseMeasure(getAttr(extent.openTag, "cx")) : undefined,
    heightEmu: extent ? parseMeasure(getAttr(extent.openTag, "cy")) : undefined,
    alt: docPr ? (getAttr(docPr.openTag, "descr") ?? getAttr(docPr.openTag, "name")) : undefined,
  };
}

function parseTable(xml: string, el: ElementSpan): Table {
  const grid: number[] = [];
  const gridEl = findElement(xml, "w:tblGrid", { from: el.innerStart, to: el.innerEnd });
  if (gridEl) {
    for (const gc of childElements(xml, gridEl.innerStart, gridEl.innerEnd)) {
      if (localName(gc.name) === "gridCol") grid.push(parseMeasure(getAttr(gc.openTag, "w:w")));
    }
  }
  const rows: TableRow[] = [];
  for (const tr of childElements(xml, el.innerStart, el.innerEnd)) {
    if (localName(tr.name) === "tr") rows.push(parseRow(xml, tr));
  }
  return { type: "table", grid, rows, source: { start: el.outerStart, end: el.outerEnd } };
}

function parseRow(xml: string, el: ElementSpan): TableRow {
  const trPr = findElement(xml, "w:trPr", { from: el.innerStart, to: el.innerEnd });
  const isHeader = trPr ? findElement(xml, "w:tblHeader", { from: trPr.innerStart, to: trPr.innerEnd }) !== undefined : false;
  const cells: TableCell[] = [];
  for (const tc of childElements(xml, el.innerStart, el.innerEnd)) {
    if (localName(tc.name) === "tc") cells.push(parseCell(xml, tc));
  }
  return { cells, isHeader, source: { start: el.outerStart, end: el.outerEnd } };
}

function parseCell(xml: string, el: ElementSpan): TableCell {
  const props: TableCellProps = {};
  const tcPr = findElement(xml, "w:tcPr", { from: el.innerStart, to: el.innerEnd });
  if (tcPr) {
    const gs = findElement(xml, "w:gridSpan", { from: tcPr.innerStart, to: tcPr.innerEnd });
    if (gs) props.gridSpan = parseMeasure(getAttr(gs.openTag, "w:val"), 1);
    const vm = findElement(xml, "w:vMerge", { from: tcPr.innerStart, to: tcPr.innerEnd });
    if (vm) props.vMerge = getAttr(vm.openTag, "w:val") === "restart" ? "restart" : "continue";
    const tcW = findElement(xml, "w:tcW", { from: tcPr.innerStart, to: tcPr.innerEnd });
    if (tcW) props.width = { value: parseMeasure(getAttr(tcW.openTag, "w:w")), type: getAttr(tcW.openTag, "w:type") ?? "dxa" };
  }
  return { props, blocks: parseBlocks(xml, el.innerStart, el.innerEnd), source: { start: el.outerStart, end: el.outerEnd } };
}

function parseSectPr(xml: string, el: ElementSpan): SectionProps {
  const s: SectionProps = {};
  const at = (name: string): ElementSpan | undefined => findElement(xml, name, { from: el.innerStart, to: el.innerEnd });

  const pgSz = at("w:pgSz");
  if (pgSz) {
    s.pageSize = {
      width: parseMeasure(getAttr(pgSz.openTag, "w:w")),
      height: parseMeasure(getAttr(pgSz.openTag, "w:h")),
      orient: getAttr(pgSz.openTag, "w:orient"),
    };
  }

  const pgMar = at("w:pgMar");
  if (pgMar) {
    const m = (a: string): number | undefined => {
      const v = getAttr(pgMar.openTag, a);
      return v !== undefined ? parseMeasure(v) : undefined;
    };
    s.margins = { top: m("w:top"), right: m("w:right"), bottom: m("w:bottom"), left: m("w:left"), header: m("w:header"), footer: m("w:footer"), gutter: m("w:gutter") };
  }

  const cols = at("w:cols");
  if (cols) {
    s.cols = { num: parseMeasure(getAttr(cols.openTag, "w:num"), 1), space: getAttr(cols.openTag, "w:space") !== undefined ? parseMeasure(getAttr(cols.openTag, "w:space")) : undefined };
  }

  const headerRefs: { type: string; rId: string }[] = [];
  const footerRefs: { type: string; rId: string }[] = [];
  for (const ref of childElements(xml, el.innerStart, el.innerEnd)) {
    const ln = localName(ref.name);
    const rId = getAttr(ref.openTag, "r:id");
    if (!rId) continue;
    if (ln === "headerReference") headerRefs.push({ type: getAttr(ref.openTag, "w:type") ?? "default", rId });
    else if (ln === "footerReference") footerRefs.push({ type: getAttr(ref.openTag, "w:type") ?? "default", rId });
  }
  if (headerRefs.length) s.headerRefs = headerRefs;
  if (footerRefs.length) s.footerRefs = footerRefs;

  if (at("w:titlePg")) s.titlePage = true;
  return s;
}

// ── convenience accessors ─────────────────────────────────────────────────────────────────────

/** Flatten a paragraph's inline text (runs + hyperlink runs). */
export function paragraphText(p: Paragraph): string {
  const walk = (nodes: Inline[]): string =>
    nodes.map((n) => (n.type === "run" ? n.text : n.type === "hyperlink" ? walk(n.children) : "")).join("");
  return walk(p.children);
}

/** All paragraph text in the body, in order (descends into tables). */
export function documentText(model: DocumentModel): string {
  const fromBlocks = (blocks: Block[]): string[] =>
    blocks.flatMap((b) =>
      b.type === "paragraph"
        ? [paragraphText(b)]
        : b.rows.flatMap((r) => r.cells.flatMap((c) => fromBlocks(c.blocks))),
    );
  return fromBlocks(model.body).join("\n");
}
