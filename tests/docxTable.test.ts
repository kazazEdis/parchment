import { describe, it, expect } from "vitest";
import { resolveTableGrid } from "../src/table";
import type { Table, TableCell } from "../src/model";

const S = { start: 0, end: 0 };
const cell = (props: TableCell["props"] = {}): TableCell => ({ props, blocks: [], source: S });
const row = (cells: TableCell[]) => ({ cells, isHeader: false, source: S });

describe("table: span resolution", () => {
  it("folds vMerge continue cells into the restart cell's rowSpan", () => {
    const table: Table = {
      type: "table", grid: [100, 100, 100], source: S,
      rows: [
        row([cell(), cell({ vMerge: "restart" }), cell()]),
        row([cell(), cell({ vMerge: "continue" }), cell()]),
        row([cell(), cell({ vMerge: "continue" }), cell()]),
      ],
    };
    const g = resolveTableGrid(table);
    expect(g.map((r) => r.length)).toEqual([3, 2, 2]); // merged cells dropped from rows 2,3
    expect(g[0][1].rowSpan).toBe(3); // B spans all three rows
  });

  it("maps gridSpan to colSpan", () => {
    const table: Table = {
      type: "table", grid: [100, 100, 100], source: S,
      rows: [row([cell({ gridSpan: 2 }), cell()]), row([cell(), cell(), cell()])],
    };
    const g = resolveTableGrid(table);
    expect(g[0].map((c) => c.colSpan)).toEqual([2, 1]);
    expect(g[1]).toHaveLength(3);
  });

  it("infers colSpan from cell width when w:gridSpan is absent", () => {
    // 6-col grid; a data row whose 2nd cell is as wide as cols 2+3 (no gridSpan) must occupy 2 cols
    // so the trailing cells line up under the header (which DOES set gridSpan=2). Mirrors the offer
    // items table where the description column is widened by width, not span.
    const table: Table = {
      type: "table", grid: [1008, 3387, 1984, 1134, 1276, 1276], source: S,
      rows: [
        row([cell(), cell({ gridSpan: 2 }), cell(), cell(), cell()]),               // header
        row([cell({ width: { value: 1008, type: "dxa" } }),
             cell({ width: { value: 5371, type: "dxa" } }),                          // 3387+1984
             cell({ width: { value: 1134, type: "dxa" } }),
             cell({ width: { value: 1276, type: "dxa" } }),
             cell({ width: { value: 1276, type: "dxa" } })]),
      ],
    };
    const g = resolveTableGrid(table);
    expect(g[0].map((c) => c.colSpan)).toEqual([1, 2, 1, 1, 1]); // header
    expect(g[1].map((c) => c.colSpan)).toEqual([1, 2, 1, 1, 1]); // data row now aligned
  });

  it("a restart immediately followed by a non-merge clears the span", () => {
    const table: Table = {
      type: "table", grid: [100], source: S,
      rows: [row([cell({ vMerge: "restart" })]), row([cell()]), row([cell({ vMerge: "continue" })])],
    };
    const g = resolveTableGrid(table);
    expect(g[0][0].rowSpan).toBe(1); // restart not continued (row 2 is a normal cell)
    expect(g[2]).toHaveLength(1); // the later continue has no open span → rendered as its own cell
  });
});

