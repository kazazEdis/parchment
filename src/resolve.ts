// The style-engine: compose the document model with styles + numbering into *effective* properties,
// and run the numbering counter pass (ARCHITECTURE.md §1.1 "resolve late", §3.5, §3.6).
//
// Effective run props cascade:  docDefaults → paragraph-style rPr → numbering-level rPr →
//                               character-style rPr → direct run rPr
// Effective paragraph props:    docDefaults → paragraph-style pPr → numbering-level pPr → direct pPr
//
// docDefaults is applied exactly once (we use the chain-only style builders), so a character style
// never resets a paragraph style's size via its own inherited default.
import { type StyleSheet, paragraphStyleChain, runStyleChain, resolveThemeFont } from "./styles";
import { type Numbering, getLevel, formatMarker } from "./numbering";
import { type RunProps, type ParagraphProps, mergeRunProps, mergeParagraphProps } from "./props";
import { type DocumentModel, type Block, type Paragraph } from "./model";

/**
 * Effective paragraph props: docDefaults → table style → paragraph style → numbering → direct.
 * `tableStyleId` is the style of the enclosing table (when this paragraph is inside a cell); its
 * pPr applies to every paragraph in the table (e.g. TableGrid resets spacing-after to 0 and line to
 * single) and sits BELOW the paragraph's own style so a pStyle still wins. Word applies this; without
 * it cell paragraphs wrongly inherit the document's docDefaults spacing and tables render too tall.
 */
export function effectiveParagraphProps(
  sheet: StyleSheet,
  numbering: Numbering | undefined,
  pPr: ParagraphProps,
  tableStyleId?: string,
): ParagraphProps {
  let out = sheet.docDefaults.pPr;
  if (tableStyleId) out = mergeParagraphProps(out, paragraphStyleChain(sheet, tableStyleId).pPr);
  out = mergeParagraphProps(out, paragraphStyleChain(sheet, pPr.styleId).pPr);
  if (numbering && pPr.numbering) {
    const lvl = getLevel(numbering, pPr.numbering.numId, pPr.numbering.level);
    if (lvl) out = mergeParagraphProps(out, lvl.pPr);
  }
  return mergeParagraphProps(out, pPr);
}

/**
 * Effective run props for a run inside a paragraph: docDefaults → paragraph-style rPr → numbering-
 * level rPr → character-style rPr → direct rPr. `paragraphStyleId` is the run's paragraph's style.
 */
export function effectiveRunProps(
  sheet: StyleSheet,
  numbering: Numbering | undefined,
  paragraphPPr: ParagraphProps,
  runRPr: RunProps,
  tableStyleId?: string,
): RunProps {
  let out = sheet.docDefaults.rPr;
  if (tableStyleId) out = mergeRunProps(out, paragraphStyleChain(sheet, tableStyleId).rPr);
  out = mergeRunProps(out, paragraphStyleChain(sheet, paragraphPPr.styleId).rPr);
  if (numbering && paragraphPPr.numbering) {
    const lvl = getLevel(numbering, paragraphPPr.numbering.numId, paragraphPPr.numbering.level);
    if (lvl) out = mergeRunProps(out, lvl.rPr);
  }
  if (runRPr.styleId) out = mergeRunProps(out, runStyleChain(sheet, runRPr.styleId));
  return substituteThemeFonts(sheet, mergeRunProps(out, runRPr));
}

/** After the cascade: a theme font token (w:asciiTheme="minorHAnsi") with no concrete w:ascii at
 *  the same-or-later level resolves to the theme part's typeface. Without this, theme-driven runs
 *  (Word's default — docDefaults uses minorHAnsi) had no font-family at all and fell back to the
 *  browser's serif default, breaking metrics-sensitive layout. */
function substituteThemeFonts(sheet: StyleSheet, p: RunProps): RunProps {
  const f = p.fonts;
  if (!f || !sheet.themeFonts) return p;
  const ascii = f.ascii ?? resolveThemeFont(sheet.themeFonts, f.asciiTheme);
  const hAnsi = f.hAnsi ?? resolveThemeFont(sheet.themeFonts, f.hAnsiTheme);
  if (ascii === f.ascii && hAnsi === f.hAnsi) return p;
  return { ...p, fonts: { ...f, ascii, hAnsi } };
}

/** The marker run props for a list paragraph's bullet/number (numbering-level rPr over the cascade). */
export function markerRunProps(
  sheet: StyleSheet,
  numbering: Numbering,
  pPr: ParagraphProps,
): RunProps {
  let out = mergeRunProps(sheet.docDefaults.rPr, paragraphStyleChain(sheet, pPr.styleId).rPr);
  if (pPr.numbering) {
    const lvl = getLevel(numbering, pPr.numbering.numId, pPr.numbering.level);
    if (lvl) out = mergeRunProps(out, lvl.rPr);
  }
  return substituteThemeFonts(sheet, out);
}

/**
 * Walk the document in reading order and assign each list paragraph its rendered marker string
 * (e.g. "1.", "a)", "•"). Counters are kept per w:numId and per level; entering a level starts it at
 * the level's `start` (first time) or increments it, and resets all deeper levels — the standard
 * Word restart behaviour. Returns a map keyed by the Paragraph node identity.
 */
export function assignListNumbers(model: DocumentModel, numbering: Numbering): Map<Paragraph, string> {
  const result = new Map<Paragraph, string>();
  const counters = new Map<number, number[]>(); // numId → current value per level
  const started = new Map<number, Set<number>>(); // numId → levels already begun

  const visitBlocks = (blocks: Block[]): void => {
    for (const b of blocks) {
      if (b.type === "paragraph") visitParagraph(b);
      else for (const row of b.rows) for (const cell of row.cells) visitBlocks(cell.blocks);
    }
  };

  const visitParagraph = (p: Paragraph): void => {
    const num = p.pPr.numbering;
    if (!num) return;
    const level = getLevel(numbering, num.numId, num.level);
    if (!level) return; // references a non-existent list — not numbered

    const arr = counters.get(num.numId) ?? [];
    const begun = started.get(num.numId) ?? new Set<number>();

    if (begun.has(num.level)) arr[num.level] = (arr[num.level] ?? 0) + 1;
    else {
      arr[num.level] = level.start;
      begun.add(num.level);
    }
    // Entering a level restarts everything deeper.
    for (let k = num.level + 1; k < arr.length; k++) {
      arr[k] = 0;
      begun.delete(k);
    }

    counters.set(num.numId, arr);
    started.set(num.numId, begun);
    result.set(p, formatMarker(numbering, num.numId, num.level, arr));
  };

  visitBlocks(model.body);
  return result;
}
