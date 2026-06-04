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

