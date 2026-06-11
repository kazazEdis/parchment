// Shared table-cell CSS resolution (ARCHITECTURE.md §3.7). Used by BOTH the print/preview renderer
// (PaginatedDocxView) and the interactive editor (DocxEditor) so a table renders identically in each
// — previously each path had its own copy (the editor drew a hardcoded 1px grid on every cell),
// which is exactly how the editor canvas drifted away from the faithful preview.
import type { CSSProperties } from "react";
import type { BorderSide, Borders } from "./model";

/** One OOXML border edge → a CSS border value, or undefined for "no edge". val "none"/"nil" (or an
 *  absent side) means Word draws nothing — a borderless table must show NO rule (drawing a 1px rule on
 *  every cell turns it into a heavy grid). sz is in eighths of a point. */
export function borderCss(s?: BorderSide): string | undefined {
  if (!s || s.val === "none" || s.val === "nil") return undefined;
  const px = Math.max(1, Math.round((s.sz / 8) * (96 / 72)));
  const color = !s.color || s.color === "auto" ? "#000" : `#${s.color}`;
  return `${px}px solid ${color}`;
}

/** Resolve a cell's four border edges in Word's cascade: the cell's own w:tcBorders (cbd) override the
 *  table's borders (tb), where a side is the outer edge (top/bottom/left/right) on the table's
 *  first/last row/column, else the interior rule (insideH between rows, insideV between columns). */
export function cellBorderStyle(
  tb: Borders | undefined,
  cbd: Borders | undefined,
  pos: { firstRow: boolean; lastRow: boolean; firstCol: boolean; lastCol: boolean },
): Pick<CSSProperties, "borderTop" | "borderRight" | "borderBottom" | "borderLeft"> {
  return {
    borderTop: borderCss(cbd?.top ?? (pos.firstRow ? tb?.top : tb?.insideH)),
    borderBottom: borderCss(cbd?.bottom ?? (pos.lastRow ? tb?.bottom : tb?.insideH)),
    borderLeft: borderCss(cbd?.left ?? (pos.firstCol ? tb?.left : tb?.insideV)),
    borderRight: borderCss(cbd?.right ?? (pos.lastCol ? tb?.right : tb?.insideV)),
  };
}
