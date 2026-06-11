import { describe, it, expect } from "vitest";
import { borderCss, cellBorderStyle } from "../src/tableCss";
import type { Borders } from "../src/model";

// Shared by the print/preview renderer AND the interactive editor — the regression that made the
// editor canvas draw a heavy grid on borderless tables was a private, drifted copy of this logic.

describe("borderCss", () => {
  it("returns undefined for an absent side or a none/nil edge", () => {
    expect(borderCss(undefined)).toBeUndefined();
    expect(borderCss({ val: "none", sz: 4, color: "auto" })).toBeUndefined();
    expect(borderCss({ val: "nil", sz: 0, color: "auto" })).toBeUndefined();
  });

  it("converts sz (eighths of a pt) to >=1px and resolves the colour", () => {
    expect(borderCss({ val: "single", sz: 8, color: "FF0000" })).toBe("1px solid #FF0000");
    expect(borderCss({ val: "single", sz: 4, color: "auto" })).toBe("1px solid #000"); // rounds up to 1px
    expect(borderCss({ val: "single", sz: 24, color: "auto" })).toBe("4px solid #000");
    expect(borderCss({ val: "single", sz: 8, color: "" })).toBe("1px solid #000");
  });
});

describe("cellBorderStyle", () => {
  it("a borderless table (no table/cell borders) draws NOTHING — the core fix", () => {
    const s = cellBorderStyle(undefined, undefined, { firstRow: true, lastRow: false, firstCol: true, lastCol: false });
    expect(s).toEqual({ borderTop: undefined, borderBottom: undefined, borderLeft: undefined, borderRight: undefined });
  });

  it("an interior cell uses insideH/insideV, an edge cell uses the outer side", () => {
    const tb: Borders = {
      top: { val: "single", sz: 8, color: "auto" }, bottom: { val: "single", sz: 8, color: "auto" },
      left: { val: "single", sz: 8, color: "auto" }, right: { val: "single", sz: 8, color: "auto" },
      insideH: { val: "single", sz: 4, color: "AAAAAA" }, insideV: { val: "single", sz: 4, color: "AAAAAA" },
    };
    const interior = cellBorderStyle(tb, undefined, { firstRow: false, lastRow: false, firstCol: false, lastCol: false });
    expect(interior).toEqual({ borderTop: "1px solid #AAAAAA", borderBottom: "1px solid #AAAAAA", borderLeft: "1px solid #AAAAAA", borderRight: "1px solid #AAAAAA" });
    const corner = cellBorderStyle(tb, undefined, { firstRow: true, lastRow: false, firstCol: true, lastCol: false });
    expect(corner.borderTop).toBe("1px solid #000");   // outer top
    expect(corner.borderLeft).toBe("1px solid #000");  // outer left
    expect(corner.borderBottom).toBe("1px solid #AAAAAA"); // interior between rows
  });

  it("a cell's own w:tcBorders override the table border for that side", () => {
    const tb: Borders = { top: { val: "single", sz: 8, color: "auto" } };
    const cbd: Borders = { top: { val: "single", sz: 24, color: "FF0000" } };
    const s = cellBorderStyle(tb, cbd, { firstRow: true, lastRow: true, firstCol: true, lastCol: true });
    expect(s.borderTop).toBe("4px solid #FF0000");
  });

  it("a cell border of val=none suppresses the table's edge", () => {
    const tb: Borders = { top: { val: "single", sz: 8, color: "auto" } };
    const cbd: Borders = { top: { val: "none", sz: 0, color: "auto" } };
    const s = cellBorderStyle(tb, cbd, { firstRow: true, lastRow: false, firstCol: false, lastCol: false });
    expect(s.borderTop).toBeUndefined();
  });
});
