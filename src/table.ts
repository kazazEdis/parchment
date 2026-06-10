// Resolve OOXML table spans into an HTML-style render grid (ARCHITECTURE.md §3.7).
//
// OOXML encodes spans two ways: horizontal via w:gridSpan (a single w:tc covering N columns) and
// vertical via w:vMerge — a "restart" cell begins a vertical merge and each subsequent row repeats a
// "continue" w:tc in that column. HTML uses colSpan/rowSpan with the covered cells omitted. This
// pure pass converts the former to the latter: continue-cells are folded into the restart cell's
// rowSpan and dropped, so each output row lists only the cells to actually render.
import type { Table, TableCell } from "./model";

export interface RenderCell {
  cell: TableCell;
  colSpan: number;
  rowSpan: number;
}

// How many grid columns a cell covers when it has no explicit w:gridSpan: Word derives the span from
// the cell's width (w:tcW) — a cell as wide as N grid columns occupies N of them. Without this a wide
// cell (e.g. a description column the template widened instead of setting gridSpan) collapses to one
// column and every following cell in the row shifts left, misaligning the data under its headers.
function inferColSpan(grid: number[], col: number, cellWidthTwips: number | undefined): number {
  if (!cellWidthTwips || col >= grid.length) return 1;
  // Accumulate grid columns from `col` until their total reaches the cell's width (10-twip tolerance
  // for rounding). Stop at the first column that meets/exceeds it; never overruns the grid.
  let acc = 0;
  let n = 0;
  for (let c = col; c < grid.length; c++) {
    acc += grid[c];
    n += 1;
    if (acc >= cellWidthTwips - 10) break;
  }
  return Math.max(1, n);
}

export function resolveTableGrid(table: Table): RenderCell[][] {
  const result: RenderCell[][] = table.rows.map(() => []);
  const openSpan: (RenderCell | null)[] = []; // per grid column → the restart cell currently spanning down

  table.rows.forEach((row, r) => {
    let col = 0;
    for (const cell of row.cells) {
      const colSpan = cell.props.gridSpan
        ?? (cell.props.width?.type === "dxa" ? inferColSpan(table.grid, col, cell.props.width.value) : 1);
      if (cell.props.vMerge === "continue" && openSpan[col]) {
        openSpan[col]!.rowSpan += 1; // covered by the restart above — extend it, emit nothing
      } else {
        const rc: RenderCell = { cell, colSpan, rowSpan: 1 };
        result[r].push(rc);
        const isRestart = cell.props.vMerge === "restart";
        for (let c = col; c < col + colSpan; c++) openSpan[c] = isRestart ? rc : null;
      }
      col += colSpan;
    }
  });

  return result;
}
