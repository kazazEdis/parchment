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

export function resolveTableGrid(table: Table): RenderCell[][] {
  const result: RenderCell[][] = table.rows.map(() => []);
  const openSpan: (RenderCell | null)[] = []; // per grid column → the restart cell currently spanning down

  table.rows.forEach((row, r) => {
    let col = 0;
    for (const cell of row.cells) {
      const colSpan = cell.props.gridSpan ?? 1;
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
