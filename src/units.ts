// OOXML measurement units. Word stores lengths in several different units depending on the
// attribute; mixing them up is the single most common source of "it looks slightly wrong" bugs
// (ARCHITECTURE.md §3.2). Every conversion in the editor goes through here — pure number → number.
//
// Reference points (ECMA-376):
//   1 inch = 72 pt = 1440 twip = 914400 EMU = 96 px (at the conventional 96 dpi CSS reference)

export const TWIPS_PER_POINT = 20;
export const TWIPS_PER_INCH = 1440;
export const EMU_PER_INCH = 914400;
export const EMU_PER_POINT = 12700;
export const EMU_PER_CM = 360000;
export const CSS_DPI = 96; // CSS reference pixel density

// ── twips (1/20 pt) — margins, indents, tab stops, table/cell widths, page size ──
export const twipsToPoints = (twips: number): number => twips / TWIPS_PER_POINT;
export const pointsToTwips = (pt: number): number => pt * TWIPS_PER_POINT;
export const twipsToInches = (twips: number): number => twips / TWIPS_PER_INCH;
export const twipsToCm = (twips: number): number => (twips / TWIPS_PER_INCH) * 2.54;
export const twipsToPx = (twips: number, dpi: number = CSS_DPI): number => (twips / TWIPS_PER_INCH) * dpi;
export const pxToTwips = (px: number, dpi: number = CSS_DPI): number => (px / dpi) * TWIPS_PER_INCH;

// ── half-points — font size (w:sz val="24" = 12pt) and a few others ──
export const halfPointsToPoints = (hp: number): number => hp / 2;
export const pointsToHalfPoints = (pt: number): number => pt * 2;

// ── eighth-points — border widths (w:bdr/@w:sz, val="8" = 1pt) ──
export const eighthPointsToPoints = (ep: number): number => ep / 8;
export const pointsToEighthPoints = (pt: number): number => pt * 8;

// ── EMU — DrawingML image/shape sizes (wp:extent @cx/@cy) ──
export const emuToPoints = (emu: number): number => emu / EMU_PER_POINT;
export const emuToInches = (emu: number): number => emu / EMU_PER_INCH;
export const emuToPx = (emu: number, dpi: number = CSS_DPI): number => (emu / EMU_PER_INCH) * dpi;
export const pxToEmu = (px: number, dpi: number = CSS_DPI): number => Math.round((px / dpi) * EMU_PER_INCH);
export const emuToCm = (emu: number): number => emu / EMU_PER_CM;

// ── fiftieths-of-a-percent — w:tblW/@w:w when type="pct" (val="5000" = 100%) ──
// (OOXML transitional uses 1/50 of a percent; the newer "50%" string form is not handled here.)
export const fiftiethsToPercent = (val: number): number => val / 50;
export const percentToFiftieths = (pct: number): number => Math.round(pct * 50);

// ── line spacing — w:spacing/@w:line. With lineRule="auto" it is a multiple in 240ths
//    (240 = single, 360 = 1.5×, 480 = double). With atLeast/exactly it is twips. ──
export const lineAutoToMultiple = (line: number): number => line / 240;
export const multipleToLineAuto = (mult: number): number => Math.round(mult * 240);

/**
 * Parse an integer-ish OOXML measurement attribute. Tolerant of undefined/whitespace; returns
 * `fallback` (default 0) when absent or non-numeric. Does NOT interpret units — the caller knows
 * which unit the attribute is in and applies the matching conversion above.
 */
export function parseMeasure(value: string | undefined | null, fallback = 0): number {
  if (value == null) return fallback;
  const s = String(value).trim();
  if (s === "") return fallback; // Number("") is 0, not NaN — treat empty as absent
  const n = Number(s);
  return Number.isFinite(n) ? n : fallback;
}
