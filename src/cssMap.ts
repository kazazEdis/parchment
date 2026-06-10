// Effective OOXML properties → CSS (ARCHITECTURE.md §3.8 Tier-1 render). Pure: takes resolved
// RunProps/ParagraphProps (from resolve.ts) and returns React.CSSProperties, converting units along
// the way. No DOM/React runtime needed — the type import is erased — so this is unit-tested in Node.
import type { CSSProperties } from "react";
import type { RunProps, ParagraphProps } from "./props";
import { twipsToPx, emuToPx, lineAutoToMultiple } from "./units";

/** OOXML named highlight colours → CSS. */
const HIGHLIGHT: Record<string, string> = {
  yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff", blue: "#0000ff",
  red: "#ff0000", darkBlue: "#000080", darkCyan: "#008080", darkGreen: "#008000",
  darkMagenta: "#800080", darkRed: "#800000", darkYellow: "#808000", darkGray: "#808080",
  lightGray: "#c0c0c0", black: "#000000", white: "#ffffff", none: "transparent",
};

/** "FF0000" → "#FF0000"; "auto" → undefined (let the cascade/inherit decide). */
function hexColor(v: string | undefined): string | undefined {
  if (!v || v === "auto") return undefined;
  return /^[0-9a-fA-F]{6}$/.test(v) ? `#${v}` : v;
}

/** Metric-compatible substitutes appended after the named font, plus a generic family so a
 *  missing font degrades to the right shape (sans vs serif) instead of the browser serif default.
 *  Carlito (≈Calibri) ships embedded via FONT_FACE_CSS; the others rely on the OS. */
const FONT_FALLBACK: Record<string, string> = {
  calibri: "Carlito, sans-serif",
  cambria: "Caladea, Georgia, serif",
  arial: "Helvetica, sans-serif",
  "times new roman": "Liberation Serif, Georgia, serif",
  verdana: "DejaVu Sans, sans-serif",
  tahoma: "DejaVu Sans, sans-serif",
  georgia: "serif",
  garamond: "serif",
  "courier new": "Liberation Mono, monospace",
  consolas: "monospace",
};

/** A font name → a CSS font-family stack with metric-compatible + generic fallbacks. */
export function fontFamilyCss(name: string): string {
  const quoted = /[ ,]/.test(name) ? `"${name}"` : name;
  const fb = FONT_FALLBACK[name.toLowerCase()];
  return fb ? `${quoted}, ${fb}` : `${quoted}, sans-serif`;
}

/** Resolved run props → inline CSS for a text span. */
export function runCss(p: RunProps): CSSProperties {
  const css: CSSProperties = {};
  if (p.bold !== undefined) css.fontWeight = p.bold ? "bold" : "normal";
  if (p.italic !== undefined) css.fontStyle = p.italic ? "italic" : "normal";

  const deco: string[] = [];
  if (p.underline !== undefined && p.underline !== "none") deco.push("underline");
  if (p.strike) deco.push("line-through");
  if (deco.length) css.textDecoration = deco.join(" ");
  else if (p.underline === "none" || p.strike === false) css.textDecoration = "none";

  const color = hexColor(p.color);
  if (color) css.color = color;
  if (p.fontSize !== undefined) css.fontSize = `${p.fontSize}pt`;
  if (p.fonts?.ascii) css.fontFamily = fontFamilyCss(p.fonts.ascii);

  if (p.highlight !== undefined) css.backgroundColor = HIGHLIGHT[p.highlight] ?? p.highlight;
  else { const shd = hexColor(p.shading); if (shd) css.backgroundColor = shd; }

  if (p.caps) css.textTransform = "uppercase";
  if (p.smallCaps) css.fontVariant = "small-caps";
  if (p.vertAlign === "superscript") { css.verticalAlign = "super"; css.fontSize = css.fontSize ?? "0.8em"; }
  else if (p.vertAlign === "subscript") { css.verticalAlign = "sub"; css.fontSize = css.fontSize ?? "0.8em"; }
  return css;
}

const ALIGN: Record<string, CSSProperties["textAlign"]> = {
  left: "left", start: "left", right: "right", end: "right", center: "center",
  both: "justify", distribute: "justify",
};

// Word's "single" line spacing (w:line=240, lineRule=auto) = the font's natural line height, not the
// em (CSS line-height:1.0). That natural height is FONT-DEPENDENT and measurably different per family
// (vs LibreOffice/Word): Calibri/Carlito body ≈ 1.073, Cambria/Caladea headings ≈ 1.174. A single
// flat factor inflates one to fix the other — sans bodies bloat, or serif headers stay too tight and
// a header rule rides up over the logo. Pick the factor from the paragraph's dominant font.
const lineFactorFor = (font?: string): number => {
  const f = (font ?? "").toLowerCase();
  // Tall-metrics serif families (Word major-theme default Cambria, its Caladea twin, Georgia, Times).
  if (/cambria|caladea|georgia|times|garamond|serif/.test(f)) return 1.17;
  return 1.08; // Calibri/Carlito/Arial-class sans (measured Word Calibri body 1.073)
};

/** Resolved paragraph props → inline CSS for a block. `font` = the paragraph's dominant ascii font,
 *  used to pick the correct single-line factor for lineRule="auto" (see lineFactorFor). */
export function paragraphCss(p: ParagraphProps, font?: string): CSSProperties {
  const css: CSSProperties = {};
  if (p.alignment) css.textAlign = ALIGN[p.alignment];

  if (p.indent) {
    if (p.indent.left !== undefined) css.marginLeft = twipsToPx(p.indent.left);
    if (p.indent.right !== undefined) css.marginRight = twipsToPx(p.indent.right);
    if (p.indent.firstLine !== undefined) css.textIndent = twipsToPx(p.indent.firstLine);
    else if (p.indent.hanging !== undefined) css.textIndent = -twipsToPx(p.indent.hanging);
  }

  if (p.spacing) {
    if (p.spacing.before !== undefined) css.marginTop = twipsToPx(p.spacing.before);
    if (p.spacing.after !== undefined) css.marginBottom = twipsToPx(p.spacing.after);
    if (p.spacing.line !== undefined) {
      css.lineHeight = (p.spacing.lineRule ?? "auto") === "auto"
        // OOXML auto line is a multiple of SINGLE line spacing, and Word's single line = the font's
        // natural line metrics (~1.15× the em for Calibri/Carlito), NOT 1.0× the font size. CSS
        // `line-height: 1.0` is too tight: text measured 9.76pt/line where Word gives 11.74pt, which
        // accumulates (a 6-line block 12pt short → header rules ride up, tables ~13px under Word).
        // Scale the multiple by the natural-line factor so "single" renders like Word.
        ? lineAutoToMultiple(p.spacing.line) * lineFactorFor(font)
        : `${twipsToPx(p.spacing.line)}px`;
    }
  }

  const shd = hexColor(p.shading);
  if (shd) css.backgroundColor = shd;
  return css;
}

/** Redline styling for a tracked insertion/deletion run (underline-green / strike-red). */
export function trackCss(type: "ins" | "del"): CSSProperties {
  return type === "ins"
    ? { color: "#2e7d32", textDecoration: "underline" }
    : { color: "#c62828", textDecoration: "line-through" };
}

/** EMU width/height → CSS px sizing for an image/drawing. */
export function drawingCss(widthEmu?: number, heightEmu?: number): CSSProperties {
  const css: CSSProperties = {};
  if (widthEmu !== undefined) css.width = emuToPx(widthEmu);
  if (heightEmu !== undefined) css.height = emuToPx(heightEmu);
  return css;
}
