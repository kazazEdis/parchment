// styles.xml — style definitions and the property cascade (ARCHITECTURE.md §3.5).
//
// A run's effective formatting is resolved in order: docDefaults → paragraph style (following its
// w:basedOn chain to the root, applied root-first) → character style → numbering level → direct
// rPr/pPr. This module parses the definitions and resolves the docDefaults + basedOn portion; the
// model/render layer composes character-style and direct formatting on top (override-merge, which
// is order-sensitive but idempotent, so applying docDefaults via the paragraph style is safe).
import { findElement, findElements, getAttr, childElements, localName } from "./xml";
import {
  type RunProps,
  type ParagraphProps,
  parseRunProps,
  parseParagraphProps,
  mergeRunProps,
  mergeParagraphProps,
} from "./props";
import type { Borders, BorderSide } from "./model";

// Parse a table style's w:tblPr > w:tblBorders into a typed border set (the source of a styled
// table's grid, e.g. the built-in "TableGrid"). Mirrors model.parseBorders but works off a string.
function parseTableStyleBorders(body: string): Borders | undefined {
  const tblPr = findElement(body, "w:tblPr");
  if (!tblPr) return undefined;
  const c = findElement(body, "w:tblBorders", { from: tblPr.innerStart, to: tblPr.innerEnd });
  if (!c) return undefined;
  const b: Borders = {};
  for (const side of childElements(body, c.innerStart, c.innerEnd)) {
    const ln = localName(side.name);
    if (ln === "top" || ln === "left" || ln === "bottom" || ln === "right" || ln === "insideH" || ln === "insideV") {
      (b as Record<string, BorderSide>)[ln] = {
        val: getAttr(side.openTag, "w:val") ?? "single",
        sz: Number(getAttr(side.openTag, "w:sz") ?? 0) || 0,
        color: getAttr(side.openTag, "w:color") ?? "auto",
      };
    }
  }
  return Object.keys(b).length ? b : undefined;
}

export type StyleType = "paragraph" | "character" | "table" | "numbering";

export interface StyleDef {
  styleId: string;
  type: StyleType;
  name?: string;
  basedOn?: string;
  next?: string;
  link?: string;
  isDefault: boolean;
  /** This style's own (unresolved) paragraph props. */
  pPr: ParagraphProps;
  /** This style's own (unresolved) run props. */
  rPr: RunProps;
  /** This style's own table borders (type="table" styles only), e.g. "TableGrid". */
  tblBorders?: Borders;
}

export interface StyleSheet {
  styles: Map<string, StyleDef>;
  docDefaults: { pPr: ParagraphProps; rPr: RunProps };
  defaultParagraphStyleId?: string;
  defaultCharacterStyleId?: string;
  /** Latin typefaces from the theme part (a:majorFont / a:minorFont). Resolves theme font tokens
   *  (w:asciiTheme="minorHAnsi" → minor, "majorHAnsi" → major) in the run cascade. */
  themeFonts?: { major?: string; minor?: string };
}

/** Parse theme1.xml's major/minor latin typefaces (what w:asciiTheme tokens point at). */
export function parseThemeFonts(themeXml: string | undefined): { major?: string; minor?: string } | undefined {
  if (!themeXml) return undefined;
  const grab = (tag: string): string | undefined => {
    const scheme = findElement(themeXml, tag);
    if (!scheme) return undefined;
    const latin = findElement(themeXml, "a:latin", { from: scheme.innerStart, to: scheme.innerEnd });
    const tf = latin ? getAttr(latin.openTag, "typeface") : undefined;
    return tf || undefined;
  };
  const major = grab("a:majorFont");
  const minor = grab("a:minorFont");
  return major || minor ? { major, minor } : undefined;
}

/** Map a theme font token to the concrete typeface from the theme part. */
export function resolveThemeFont(themeFonts: { major?: string; minor?: string } | undefined, token: string | undefined): string | undefined {
  if (!themeFonts || !token) return undefined;
  return token.startsWith("major") ? themeFonts.major : token.startsWith("minor") ? themeFonts.minor : undefined;
}

const innerOf = (xml: string, name: string, from = 0): string | undefined => {
  const el = findElement(xml, name, { from });
  return el ? xml.slice(el.innerStart, el.innerEnd) : undefined;
};

/** Parse styles.xml into a style table + docDefaults. Pass the theme part's XML so theme font
 *  tokens (w:asciiTheme) resolve to concrete typefaces in the cascade; omitted → tokens ignored. */
export function parseStyles(stylesXml: string, themeXml?: string): StyleSheet {
  const styles = new Map<string, StyleDef>();

  // docDefaults: w:docDefaults > {w:rPrDefault > w:rPr, w:pPrDefault > w:pPr}
  let defR: RunProps = {};
  let defP: ParagraphProps = {};
  const dd = findElement(stylesXml, "w:docDefaults");
  if (dd) {
    const ddInner = stylesXml.slice(dd.innerStart, dd.innerEnd);
    const rd = innerOf(ddInner, "w:rPrDefault");
    if (rd !== undefined) { const r = innerOf(rd, "w:rPr"); if (r !== undefined) defR = parseRunProps(r); }
    const pd = innerOf(ddInner, "w:pPrDefault");
    if (pd !== undefined) { const p = innerOf(pd, "w:pPr"); if (p !== undefined) defP = parseParagraphProps(p); }
  }

  let defaultParagraphStyleId: string | undefined;
  let defaultCharacterStyleId: string | undefined;

  for (const span of findElements(stylesXml, "w:style")) {
    const styleId = getAttr(span.openTag, "w:styleId");
    if (!styleId) continue;
    const type = (getAttr(span.openTag, "w:type") ?? "paragraph") as StyleType;
    const isDefault = getAttr(span.openTag, "w:default") === "1" || getAttr(span.openTag, "w:default") === "true";
    const body = stylesXml.slice(span.innerStart, span.innerEnd);

    // pPr is a direct child; the style-level rPr follows it (a w:rPr nested inside pPr is the
    // paragraph-mark props and is consumed by parseParagraphProps, not the style rPr).
    const pPrEl = findElement(body, "w:pPr");
    const pPr = pPrEl ? parseParagraphProps(body.slice(pPrEl.innerStart, pPrEl.innerEnd)) : {};
    const rPrEl = findElement(body, "w:rPr", { from: pPrEl ? pPrEl.outerEnd : 0 });
    const rPr = rPrEl ? parseRunProps(body.slice(rPrEl.innerStart, rPrEl.innerEnd)) : {};

    const def: StyleDef = {
      styleId,
      type,
      name: getAttr(findElement(body, "w:name")?.openTag ?? "", "w:val"),
      basedOn: getAttr(findElement(body, "w:basedOn")?.openTag ?? "", "w:val"),
      next: getAttr(findElement(body, "w:next")?.openTag ?? "", "w:val"),
      link: getAttr(findElement(body, "w:link")?.openTag ?? "", "w:val"),
      isDefault,
      pPr,
      rPr,
      tblBorders: type === "table" ? parseTableStyleBorders(body) : undefined,
    };
    styles.set(styleId, def);
    if (isDefault && type === "paragraph") defaultParagraphStyleId = styleId;
    if (isDefault && type === "character") defaultCharacterStyleId = styleId;
  }

  return { styles, docDefaults: { pPr: defP, rPr: defR }, defaultParagraphStyleId, defaultCharacterStyleId, themeFonts: parseThemeFonts(themeXml) };
}

/** Follow w:basedOn from `styleId` up to the root, returning ids root-first. Cycle-guarded. */
function basedOnChain(sheet: StyleSheet, styleId: string): string[] {
  const chain: string[] = [];
  const seen = new Set<string>();
  let cur: string | undefined = styleId;
  while (cur && !seen.has(cur)) {
    const def = sheet.styles.get(cur);
    if (!def) break;
    seen.add(cur);
    chain.push(cur);
    cur = def.basedOn;
  }
  return chain.reverse(); // root-first
}

/**
 * A style's own props from its w:basedOn chain, **excluding** docDefaults. Use this when composing
 * the full cascade (docDefaults → paragraph style → character style → direct) so docDefaults is
 * applied exactly once — folding it into each style would let a character style's inherited
 * docDefault size silently reset the paragraph style's size.
 */
export function paragraphStyleChain(sheet: StyleSheet, styleId: string | undefined): { pPr: ParagraphProps; rPr: RunProps } {
  let pPr: ParagraphProps = {};
  let rPr: RunProps = {};
  if (!styleId) return { pPr, rPr };
  for (const id of basedOnChain(sheet, styleId)) {
    const def = sheet.styles.get(id);
    if (!def) continue;
    pPr = mergeParagraphProps(pPr, def.pPr);
    rPr = mergeRunProps(rPr, def.rPr);
  }
  return { pPr, rPr };
}

/** A character style's own run props from its basedOn chain, excluding docDefaults. */
export function runStyleChain(sheet: StyleSheet, styleId: string | undefined): RunProps {
  let rPr: RunProps = {};
  if (!styleId) return rPr;
  for (const id of basedOnChain(sheet, styleId)) {
    const def = sheet.styles.get(id);
    if (def) rPr = mergeRunProps(rPr, def.rPr);
  }
  return rPr;
}

/** Resolve a paragraph style to effective {pPr, rPr}: docDefaults + basedOn chain. */
export function resolveParagraphStyle(sheet: StyleSheet, styleId: string | undefined): { pPr: ParagraphProps; rPr: RunProps } {
  const chain = paragraphStyleChain(sheet, styleId);
  return {
    pPr: mergeParagraphProps(sheet.docDefaults.pPr, chain.pPr),
    rPr: mergeRunProps(sheet.docDefaults.rPr, chain.rPr),
  };
}

/** Resolve a character style to effective run props: docDefaults + basedOn chain. */
export function resolveRunStyle(sheet: StyleSheet, styleId: string | undefined): RunProps {
  return mergeRunProps(sheet.docDefaults.rPr, runStyleChain(sheet, styleId));
}

/** A table style's effective borders (its w:basedOn chain, root-first, per-side override). Undefined
 *  when the style defines none — the renderer then falls back to the table's own inline borders. */
export function resolveTableStyleBorders(sheet: StyleSheet, styleId: string | undefined): Borders | undefined {
  if (!styleId) return undefined;
  let out: Borders | undefined;
  for (const id of basedOnChain(sheet, styleId)) {
    const def = sheet.styles.get(id);
    if (def?.tblBorders) out = { ...out, ...def.tblBorders };
  }
  return out;
}
