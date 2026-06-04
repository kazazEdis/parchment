import { describe, it, expect } from "vitest";
import { packPages, computePageBreaks } from "../src/paginate";

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
});

