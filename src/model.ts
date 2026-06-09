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
  /** This run is the result of a PAGE/NUMPAGES field — the paginated view substitutes the live
   *  current-page / total-page number per page (the parsed `text` is the stale cached value). */
  field?: "PAGE" | "NUMPAGES";
}

export interface Drawing {
  type: "drawing";
  anchored: boolean;
  /** Relationship id of the image blip (r:embed), resolved to bytes via document.xml.rels. */
  rEmbed?: string;
  widthEmu?: number;
  heightEmu?: number;
  alt?: string;
  /** For an anchored (floating) drawing: horizontal/vertical offset in EMU from its anchor
   *  origin (posOffset). The renderer absolute-positions the image at these offsets relative to
   *  its paragraph; `behindDoc` drawings sit behind the text. align-based anchors → undefined. */
  anchorXEmu?: number;
  anchorYEmu?: number;
  behindDoc?: boolean;
  /** a:srcRect — fraction (0..1) cropped off each edge of the source image before it's scaled into
   *  the extent box. Word uses this to show only a sub-region (e.g. a logo cut from a screenshot). */
  crop?: { l: number; t: number; r: number; b: number };
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

// One table/cell border edge. `val` is the OOXML line style ("single", "none"/"nil", "dashed"…);
// `sz` is in eighths of a point; `color` is a hex RRGGBB or "auto".
export interface BorderSide { val: string; sz: number; color: string }
// A border set as it appears on w:tblBorders (table) or w:tcBorders (cell). insideH/insideV apply
// only to table-level borders (the rules BETWEEN cells).
export interface Borders { top?: BorderSide; left?: BorderSide; bottom?: BorderSide; right?: BorderSide; insideH?: BorderSide; insideV?: BorderSide }

export interface TableCellProps {
  gridSpan?: number;
  vMerge?: "restart" | "continue";
  width?: { value: number; type: string };
  borders?: Borders;     // w:tcBorders — per-side overrides of the table border
  shd?: string;          // w:shd w:fill — cell background as hex RRGGBB (no "#"); "auto"/none → undefined
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
  borders?: Borders;     // w:tblBorders — table-level edges + insideH/insideV rules (inline)
  styleId?: string;      // w:tblStyle — borders may come from this style (e.g. "TableGrid")
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
    else if (ln === "sdt") {
      // Block-level structured-document-tag (content control): unwrap its w:sdtContent and parse
      // the blocks inside (Word commonly wraps a header/footer's content in one).
      const content = findElement(xml, "w:sdtContent", { from: el.innerStart, to: el.innerEnd });
      if (content) for (const b of parseBlocks(xml, content.innerStart, content.innerEnd)) out.push(b);
    }
    // sectPr (handled separately), bookmarkStart/End → skipped in v1
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

// Classify a field instruction string → the field kind we substitute live (else undefined).
function fieldKind(instr: string): "PAGE" | "NUMPAGES" | undefined {
  return /NUMPAGES/i.test(instr) ? "NUMPAGES" : /\bPAGE\b/i.test(instr) ? "PAGE" : undefined;
}

function parseInlines(xml: string, from: number, to: number, track?: TrackChange): Inline[] {
  const out: Inline[] = [];
  // Complex-field state (spans sibling runs): begin → instrText → separate → RESULT runs → end.
  let fldInstr: string | null = null;   // accumulated instruction while inside a field, else null
  let fldInResult = false;              // true between `separate` and `end` (the displayed value)
  for (const el of childElements(xml, from, to)) {
    const ln = localName(el.name);
    if (ln === "r") {
      const runXml = inner(xml, el);
      const fc = findElement(runXml, "w:fldChar");
      if (fc) {
        const t = getAttr(fc.openTag, "w:fldCharType");
        if (t === "begin") { fldInstr = ""; fldInResult = false; }
        else if (t === "separate") fldInResult = true;
        else if (t === "end") { fldInstr = null; fldInResult = false; }
        continue; // field-control run has no rendered text
      }
      const instr = findElement(runXml, "w:instrText");
      if (instr && fldInstr !== null) { fldInstr += inner(runXml, instr); continue; }
      const before = out.length;
      parseRunInto(xml, el, out, track);
      const kind = fldInResult && fldInstr ? fieldKind(fldInstr) : undefined;
      if (kind) for (let i = before; i < out.length; i++) { const r = out[i]; if (r.type === "run") r.field = kind; }
    }
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
      // Unwrap wrappers + simple fields → their cached result runs render as content. A simple
      // PAGE/NUMPAGES field tags its result runs so the paginated view substitutes live numbers.
      const kids = parseInlines(xml, el.innerStart, el.innerEnd, track);
      const fk = ln === "fldSimple" ? fieldKind(getAttr(el.openTag, "w:instr") ?? "") : undefined;
      if (fk) for (const c of kids) if (c.type === "run") c.field = fk;
      for (const c of kids) out.push(c);
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

// a:srcRect l/t/r/b are ST_Percentage (100000 = 100%) cropped off each edge. Returns fractions, or
// undefined when there's no crop (all edges 0 / element absent).
function parseSrcRect(body: string): Drawing["crop"] {
  const sr = findElement(body, "a:srcRect");
  if (!sr) return undefined;
  const f = (a: string): number => (parseFloat(getAttr(sr.openTag, a) ?? "0") || 0) / 100000;
  const c = { l: f("l"), t: f("t"), r: f("r"), b: f("b") };
  return c.l || c.t || c.r || c.b ? c : undefined;
}

function parseDrawing(xml: string, el: ElementSpan): Drawing {
  const body = inner(xml, el);
  const extent = findElement(body, "wp:extent");
  const blip = findElement(body, "a:blip");
  const docPr = findElement(body, "wp:docPr");
  const anchor = findElement(body, "wp:anchor");
  // Anchored (floating) drawing: capture its posOffset (EMU) so the renderer can place it. Only
  // the posOffset form is handled (the common case for header/footer logos); align-based stays undefined.
  let anchorXEmu: number | undefined, anchorYEmu: number | undefined, behindDoc: boolean | undefined;
  if (anchor) {
    behindDoc = getAttr(anchor.openTag, "behindDoc") === "1";
    const offOf = (which: string): number | undefined => {
      const pos = findElement(body, which);
      if (!pos) return undefined;
      const sub = inner(body, pos);
      const o = findElement(sub, "wp:posOffset");
      return o ? parseMeasure(inner(sub, o).trim()) : undefined;
    };
    anchorXEmu = offOf("wp:positionH");
    anchorYEmu = offOf("wp:positionV");
  }
  return {
    type: "drawing",
    anchored: anchor !== undefined,
    rEmbed: blip ? getAttr(blip.openTag, "r:embed") : undefined,
    widthEmu: extent ? parseMeasure(getAttr(extent.openTag, "cx")) : undefined,
    heightEmu: extent ? parseMeasure(getAttr(extent.openTag, "cy")) : undefined,
    alt: docPr ? (getAttr(docPr.openTag, "descr") ?? getAttr(docPr.openTag, "name")) : undefined,
    anchorXEmu, anchorYEmu, behindDoc,
    crop: parseSrcRect(body),
  };
}

// Parse a w:tblBorders / w:tcBorders container (inside the given properties element) into a typed
// Borders set. Returns undefined when the container is absent or empty.
function parseBorders(xml: string, prEl: ElementSpan, container: "w:tblBorders" | "w:tcBorders"): Borders | undefined {
  const c = findElement(xml, container, { from: prEl.innerStart, to: prEl.innerEnd });
  if (!c) return undefined;
  const b: Borders = {};
  for (const side of childElements(xml, c.innerStart, c.innerEnd)) {
    const ln = localName(side.name);
    if (ln === "top" || ln === "left" || ln === "bottom" || ln === "right" || ln === "insideH" || ln === "insideV") {
      (b as Record<string, BorderSide>)[ln] = {
        val: getAttr(side.openTag, "w:val") ?? "single",
        sz: parseMeasure(getAttr(side.openTag, "w:sz"), 0),
        color: getAttr(side.openTag, "w:color") ?? "auto",
      };
    }
  }
  return Object.keys(b).length ? b : undefined;
}

function parseTable(xml: string, el: ElementSpan): Table {
  const grid: number[] = [];
  const gridEl = findElement(xml, "w:tblGrid", { from: el.innerStart, to: el.innerEnd });
  if (gridEl) {
    for (const gc of childElements(xml, gridEl.innerStart, gridEl.innerEnd)) {
      if (localName(gc.name) === "gridCol") grid.push(parseMeasure(getAttr(gc.openTag, "w:w")));
    }
  }
  // Table-level borders live in w:tblPr. Scope the search to tblPr so a cell's w:tcBorders
  // (which also sit under this w:tbl) can't be mistaken for the table's.
  let borders: Borders | undefined;
  let styleId: string | undefined;
  const tblPr = findElement(xml, "w:tblPr", { from: el.innerStart, to: el.innerEnd });
  if (tblPr) {
    borders = parseBorders(xml, tblPr, "w:tblBorders");
    const ts = findElement(xml, "w:tblStyle", { from: tblPr.innerStart, to: tblPr.innerEnd });
    if (ts) styleId = getAttr(ts.openTag, "w:val") ?? undefined;
  }
  const rows: TableRow[] = [];
  for (const tr of childElements(xml, el.innerStart, el.innerEnd)) {
    if (localName(tr.name) === "tr") rows.push(parseRow(xml, tr));
  }
  return { type: "table", grid, rows, borders, styleId, source: { start: el.outerStart, end: el.outerEnd } };
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
    props.borders = parseBorders(xml, tcPr, "w:tcBorders");
    const shd = findElement(xml, "w:shd", { from: tcPr.innerStart, to: tcPr.innerEnd });
    if (shd) {
      const fill = getAttr(shd.openTag, "w:fill");
      if (fill && fill.toLowerCase() !== "auto" && !/^0*$/.test(fill)) props.shd = fill;
    }
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
