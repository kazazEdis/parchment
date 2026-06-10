// Run and paragraph property bags (rPr / pPr), parsed from OOXML into typed objects that support
// override-merge for the style cascade (ARCHITECTURE.md §3.4, §3.5). We model the common ~90% of
// properties; unmodelled ones survive via preserve-and-patch (the original XML is kept verbatim and
// only edited nodes are rewritten). A property left `undefined` means "not specified at this level"
// — the cascade inherits it; an explicit value (including `false` / "none") overrides the parent.
import { findElement, getAttr } from "./xml";
import { halfPointsToPoints, parseMeasure } from "./units";

export interface Fonts {
  ascii?: string;
  hAnsi?: string;
  eastAsia?: string;
  cs?: string;
  /** Theme font token ("minorHAnsi", "majorHAnsi", "minorAscii", …) from w:asciiTheme. Resolved
   *  against the theme part's major/minor latin typefaces at render time (resolve.ts); preserved
   *  verbatim on round-trip so we never bake a concrete font into a theme-driven document. */
  asciiTheme?: string;
  hAnsiTheme?: string;
}

export interface RunProps {
  bold?: boolean;
  italic?: boolean;
  /** Underline style as written ("single", "double", … or "none" to explicitly disable). */
  underline?: string;
  strike?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  /** Hex "RRGGBB" or "auto". */
  color?: string;
  /** Named highlight ("yellow", …). */
  highlight?: string;
  /** Shading fill hex for run-level shd. */
  shading?: string;
  /** Font size in points (converted from half-points). */
  fontSize?: number;
  fonts?: Fonts;
  vertAlign?: "superscript" | "subscript" | "baseline";
  /** Character-style reference (w:rStyle). */
  styleId?: string;
}

export interface Indent {
  left?: number; // twips
  right?: number;
  firstLine?: number;
  hanging?: number;
}

export interface Spacing {
  before?: number; // twips
  after?: number;
  line?: number; // twips, or 240ths when lineRule="auto"
  lineRule?: string; // "auto" | "atLeast" | "exact"
}

export interface ParagraphProps {
  /** Paragraph-style reference (w:pStyle). */
  styleId?: string;
  alignment?: "left" | "right" | "center" | "both" | "distribute" | "start" | "end";
  indent?: Indent;
  spacing?: Spacing;
  numbering?: { numId: number; level: number };
  outlineLevel?: number;
  keepNext?: boolean;
  keepLines?: boolean;
  /** Background shading fill hex. */
  shading?: string;
  /** Run properties of the paragraph mark itself (w:pPr/w:rPr) — affects empty-paragraph height. */
  markRunProps?: RunProps;
}

// ── OOXML on/off type ───────────────────────────────────────────────────────────────────────────
// A toggle element with no @w:val means "on". val in {0,false,off} means "off".
function onOff(val: string | undefined): boolean {
  if (val === undefined) return true;
  return !(val === "0" || val === "false" || val === "off");
}

/** Read a toggle child (w:b, w:i, …): undefined if absent, else its on/off boolean. */
function toggle(inner: string, tag: string): boolean | undefined {
  const el = findElement(inner, tag);
  if (!el) return undefined;
  return onOff(getAttr(el.openTag, "w:val"));
}

/** Read a child element's w:val attribute (string), undefined if the element is absent. */
function val(inner: string, tag: string): string | undefined {
  const el = findElement(inner, tag);
  return el ? getAttr(el.openTag, "w:val") : undefined;
}

/** Parse the inner XML of a `w:rPr` element into typed run properties. */
export function parseRunProps(inner: string): RunProps {
  const p: RunProps = {};
  const bold = toggle(inner, "w:b"); if (bold !== undefined) p.bold = bold;
  const italic = toggle(inner, "w:i"); if (italic !== undefined) p.italic = italic;
  const strike = toggle(inner, "w:strike"); if (strike !== undefined) p.strike = strike;
  const caps = toggle(inner, "w:caps"); if (caps !== undefined) p.caps = caps;
  const smallCaps = toggle(inner, "w:smallCaps"); if (smallCaps !== undefined) p.smallCaps = smallCaps;

  const u = findElement(inner, "w:u");
  if (u) p.underline = getAttr(u.openTag, "w:val") ?? "single";

  const color = val(inner, "w:color"); if (color !== undefined) p.color = color;
  const highlight = val(inner, "w:highlight"); if (highlight !== undefined) p.highlight = highlight;

  const sz = findElement(inner, "w:sz");
  if (sz) {
    const hp = parseMeasure(getAttr(sz.openTag, "w:val"), NaN);
    if (Number.isFinite(hp)) p.fontSize = halfPointsToPoints(hp);
  }

  const rFonts = findElement(inner, "w:rFonts");
  if (rFonts) {
    const fonts: Fonts = {};
    const ascii = getAttr(rFonts.openTag, "w:ascii"); if (ascii) fonts.ascii = ascii;
    const hAnsi = getAttr(rFonts.openTag, "w:hAnsi"); if (hAnsi) fonts.hAnsi = hAnsi;
    const eastAsia = getAttr(rFonts.openTag, "w:eastAsia"); if (eastAsia) fonts.eastAsia = eastAsia;
    const cs = getAttr(rFonts.openTag, "w:cs"); if (cs) fonts.cs = cs;
    const asciiTheme = getAttr(rFonts.openTag, "w:asciiTheme"); if (asciiTheme) fonts.asciiTheme = asciiTheme;
    const hAnsiTheme = getAttr(rFonts.openTag, "w:hAnsiTheme"); if (hAnsiTheme) fonts.hAnsiTheme = hAnsiTheme;
    if (Object.keys(fonts).length) p.fonts = fonts;
  }

  const va = val(inner, "w:vertAlign");
  if (va === "superscript" || va === "subscript" || va === "baseline") p.vertAlign = va;

  const shd = findElement(inner, "w:shd");
  if (shd) { const fill = getAttr(shd.openTag, "w:fill"); if (fill) p.shading = fill; }

  const styleId = val(inner, "w:rStyle"); if (styleId !== undefined) p.styleId = styleId;
  return p;
}

/** Parse the inner XML of a `w:pPr` element into typed paragraph properties. */
export function parseParagraphProps(inner: string): ParagraphProps {
  const p: ParagraphProps = {};

  const styleId = val(inner, "w:pStyle"); if (styleId !== undefined) p.styleId = styleId;

  const jc = val(inner, "w:jc");
  if (jc === "left" || jc === "right" || jc === "center" || jc === "both" ||
      jc === "distribute" || jc === "start" || jc === "end") p.alignment = jc;

  const ind = findElement(inner, "w:ind");
  if (ind) {
    const indent: Indent = {};
    const left = getAttr(ind.openTag, "w:left") ?? getAttr(ind.openTag, "w:start");
    const right = getAttr(ind.openTag, "w:right") ?? getAttr(ind.openTag, "w:end");
    const firstLine = getAttr(ind.openTag, "w:firstLine");
    const hanging = getAttr(ind.openTag, "w:hanging");
    if (left !== undefined) indent.left = parseMeasure(left);
    if (right !== undefined) indent.right = parseMeasure(right);
    if (firstLine !== undefined) indent.firstLine = parseMeasure(firstLine);
    if (hanging !== undefined) indent.hanging = parseMeasure(hanging);
    if (Object.keys(indent).length) p.indent = indent;
  }

  const sp = findElement(inner, "w:spacing");
  if (sp) {
    const spacing: Spacing = {};
    const before = getAttr(sp.openTag, "w:before");
    const after = getAttr(sp.openTag, "w:after");
    const line = getAttr(sp.openTag, "w:line");
    const lineRule = getAttr(sp.openTag, "w:lineRule");
    if (before !== undefined) spacing.before = parseMeasure(before);
    if (after !== undefined) spacing.after = parseMeasure(after);
    if (line !== undefined) spacing.line = parseMeasure(line);
    if (lineRule !== undefined) spacing.lineRule = lineRule;
    if (Object.keys(spacing).length) p.spacing = spacing;
  }

  const numPr = findElement(inner, "w:numPr");
  if (numPr) {
    const body = inner.slice(numPr.innerStart, numPr.innerEnd);
    const numId = val(body, "w:numId");
    const ilvl = val(body, "w:ilvl");
    if (numId !== undefined) {
      p.numbering = { numId: parseMeasure(numId), level: parseMeasure(ilvl, 0) };
    }
  }

  const outline = val(inner, "w:outlineLvl");
  if (outline !== undefined) p.outlineLevel = parseMeasure(outline);

  const keepNext = toggle(inner, "w:keepNext"); if (keepNext !== undefined) p.keepNext = keepNext;
  const keepLines = toggle(inner, "w:keepLines"); if (keepLines !== undefined) p.keepLines = keepLines;

  const shd = findElement(inner, "w:shd");
  if (shd) { const fill = getAttr(shd.openTag, "w:fill"); if (fill) p.shading = fill; }

  // The paragraph mark's own run props (w:pPr/w:rPr) — only the FIRST rPr directly under pPr.
  const rPr = findElement(inner, "w:rPr");
  if (rPr) p.markRunProps = parseRunProps(inner.slice(rPr.innerStart, rPr.innerEnd));

  return p;
}

// ── cascade merge: `over` wins per field; nested objects shallow-merge ───────────────────────────

export function mergeRunProps(base: RunProps, over: RunProps): RunProps {
  const out: RunProps = { ...base, ...over };
  if (base.fonts || over.fonts) out.fonts = { ...base.fonts, ...over.fonts };
  return out;
}

export function mergeParagraphProps(base: ParagraphProps, over: ParagraphProps): ParagraphProps {
  const out: ParagraphProps = { ...base, ...over };
  if (base.indent || over.indent) out.indent = { ...base.indent, ...over.indent };
  if (base.spacing || over.spacing) out.spacing = { ...base.spacing, ...over.spacing };
  if (base.markRunProps || over.markRunProps) {
    out.markRunProps = mergeRunProps(base.markRunProps ?? {}, over.markRunProps ?? {});
  }
  return out;
}
