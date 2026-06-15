import { describe, it, expect } from "vitest";
import { packPages, computePageBreaks, fuseKeepNextRows } from "../src/paginate";

describe("paginate: packPages", () => {
  it("packs consecutive blocks until the page is full", () => {
    expect(packPages([40, 40, 40], 100)).toEqual([[0, 1], [2]]); // 80 fits, +40 overflows
  });

  it("gives an oversized block its own page", () => {
    expect(packPages([150, 30], 100)).toEqual([[0], [1]]);
  });

  it("keeps everything on one page when it fits", () => {
    expect(packPages([10, 20, 30], 100)).toEqual([[0, 1, 2]]);
  });

  it("returns a single empty page for no blocks", () => {
    expect(packPages([], 100)).toEqual([[]]);
  });
});

describe("paginate: computePageBreaks (line-level)", () => {
  it("breaks on line boundaries so lines are never cut", () => {
    expect(computePageBreaks([20, 40, 60, 80, 100], 50)).toEqual([0, 40, 80]);
  });

  it("forces an oversized line onto its own page", () => {
    expect(computePageBreaks([60, 120], 50)).toEqual([0, 60]);
  });

  it("keeps everything on one page when it fits", () => {
    expect(computePageBreaks([20, 40, 60], 100)).toEqual([0]);
  });

  it("handles empty input", () => {
    expect(computePageBreaks([], 100)).toEqual([0]);
  });

  it("uses a shorter first page when given firstContentHeight (tall first-page header reserve)", () => {
    // First page caps at 30 (only line 20 fits → break at 20); later pages cap at 50.
    expect(computePageBreaks([20, 40, 60, 80, 100], 50, 30)).toEqual([0, 20, 60]);
  });

  it("firstContentHeight defaults to contentHeight (no special first page)", () => {
    expect(computePageBreaks([20, 40, 60, 80, 100], 50)).toEqual([0, 40, 80]);
  });

  it("does not split a table row across pages — pulls the break to the row's top", () => {
    // A row occupies y∈[40,90] (lines at 60,80). The natural break at 60 would slice it; instead the
    // whole row moves to page 2 (break at 40).
    expect(computePageBreaks([20, 40, 60, 80, 100, 120], 70, 70, [[40, 90]])).toEqual([0, 40, 100]);
  });

  it("splits a row taller than a page (force progress, no infinite loop)", () => {
    // Row [10,200] is taller than the 70 page → splitting is unavoidable; falls back to line breaks.
    expect(computePageBreaks([50, 100, 150, 200], 70, 70, [[10, 200]])).toEqual([0, 50, 100, 150]);
  });

  it("allows a break at a row boundary (touching, not bisecting)", () => {
    // Break candidate exactly at a row's edge is fine — not strictly inside.
    expect(computePageBreaks([20, 40, 60, 80], 40, 40, [[40, 80]])).toEqual([0, 40]);
  });

  it("keeps a keepNext header glued to its next row (no orphan at page foot)", () => {
    // Header row [40,55] is keepNext, content row [55,95] follows. Fused → atomic [40,95]. A natural
    // break that would leave the header on page 1 and push only the content row over is pulled back so
    // BOTH move to page 2 together (break at 40, not 55).
    const fused = fuseKeepNextRows([
      { top: 40, bottom: 55, keep: true },   // KOMERCIJALNI UVJETI header
      { top: 55, bottom: 95, keep: false },  // first terms row
    ]);
    expect(fused).toEqual([[40, 95]]);
    // fused → break pulled back to the header top (40); without fusing it would stop at the content
    // row's top (55), orphaning the header on page 1.
    expect(computePageBreaks([20, 40, 60, 80, 100], 70, 70, fused)).toEqual([0, 40]);
    expect(computePageBreaks([20, 40, 60, 80, 100], 70, 70, [[40, 55], [55, 95]])).toEqual([0, 55]);
  });
});

describe("paginate: fuseKeepNextRows", () => {
  it("merges a keepNext row with the following row, leaves standalone rows alone", () => {
    expect(fuseKeepNextRows([
      { top: 0, bottom: 10, keep: false },
      { top: 10, bottom: 20, keep: true },
      { top: 20, bottom: 40, keep: false },
      { top: 40, bottom: 60, keep: false },
    ])).toEqual([[0, 10], [10, 40], [40, 60]]);
  });
  it("chains several consecutive keepNext rows into one range", () => {
    expect(fuseKeepNextRows([
      { top: 0, bottom: 10, keep: true },
      { top: 10, bottom: 20, keep: true },
      { top: 20, bottom: 30, keep: false },
    ])).toEqual([[0, 30]]);
  });
  it("a trailing keepNext row with no successor stays its own range", () => {
    expect(fuseKeepNextRows([{ top: 0, bottom: 10, keep: true }])).toEqual([[0, 10]]);
  });
});

