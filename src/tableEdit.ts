// Table structure editing (insert/delete rows + columns). Operates on the table's original XML in
// document.xml and splices — so every w:tblPr / row / cell property (borders, shading, widths,
// styles) is preserved verbatim (preserve-and-patch), unlike re-emitting from the sparse model. v1
// assumes simple grids (no gridSpan/vMerge interplay on the edited row/column).
import { findElement, findElements } from "./xml";
import type { Table } from "./model";

const clamp = (n: number, lo: number, hi: number): number => Math.max(lo, Math.min(n, hi));

/** Blank a row's text (keep cell structure + formatting) for a freshly inserted row. */
const blankRow = (rowXml: string): string => rowXml.replace(/(<w:t\b[^>]*>)[\s\S]*?(<\/w:t>)/g, "$1$2");

function spliceTable(documentXml: string, table: Table, newTableXml: string): string {
  return documentXml.slice(0, table.source.start) + newTableXml + documentXml.slice(table.source.end);
}

/** Insert a blank copy of row `rowIndex` directly after it. */
export function insertRowAfter(documentXml: string, table: Table, rowIndex: number): string {
  const xml = documentXml.slice(table.source.start, table.source.end);
  const rows = findElements(xml, "w:tr");
  if (!rows.length) return documentXml;
  const idx = clamp(rowIndex, 0, rows.length - 1);
  const tmpl = blankRow(xml.slice(rows[idx].outerStart, rows[idx].outerEnd));
  const at = rows[idx].outerEnd;
  return spliceTable(documentXml, table, xml.slice(0, at) + tmpl + xml.slice(at));
}

/** Delete row `rowIndex` (keeps at least one row). */
export function deleteRow(documentXml: string, table: Table, rowIndex: number): string {
  const xml = documentXml.slice(table.source.start, table.source.end);
  const rows = findElements(xml, "w:tr");
  if (rows.length <= 1) return documentXml;
  const idx = clamp(rowIndex, 0, rows.length - 1);
  return spliceTable(documentXml, table, xml.slice(0, rows[idx].outerStart) + xml.slice(rows[idx].outerEnd));
}

/** Append a blank column: a w:gridCol + a blank w:tc on every row. */
export function appendColumn(documentXml: string, table: Table): string {
  let xml = documentXml.slice(table.source.start, table.source.end);
  const grid = findElement(xml, "w:tblGrid");
  const w = table.grid.length ? Math.round(table.grid.reduce((a, b) => a + b, 0) / table.grid.length) : 1000;
  if (grid) xml = xml.slice(0, grid.innerEnd) + `<w:gridCol w:w="${w}"/>` + xml.slice(grid.innerEnd);
  // append a cell before each </w:tr> (recompute rows after the grid edit shifted offsets)
  const rows = findElements(xml, "w:tr");
  for (let i = rows.length - 1; i >= 0; i--) {
    const at = rows[i].innerEnd;
    xml = xml.slice(0, at) + `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr><w:p/></w:tc>` + xml.slice(at);
  }
  return spliceTable(documentXml, table, xml);
}

/** Delete column `colIndex`: its gridCol + the matching w:tc on each row (ignores gridSpan in v1). */
export function deleteColumn(documentXml: string, table: Table, colIndex: number): string {
  let xml = documentXml.slice(table.source.start, table.source.end);
  if (table.grid.length <= 1) return documentXml;
  const idx = clamp(colIndex, 0, table.grid.length - 1);
  // remove cells right-to-left across rows first (offsets stable for the grid which precedes rows)
  const rows = findElements(xml, "w:tr");
  for (let i = rows.length - 1; i >= 0; i--) {
    const rowXml = xml.slice(rows[i].outerStart, rows[i].outerEnd);
    const cells = findElements(rowXml, "w:tc");
    if (idx >= cells.length) continue;
    const cStart = rows[i].outerStart + cells[idx].outerStart;
    const cEnd = rows[i].outerStart + cells[idx].outerEnd;
    xml = xml.slice(0, cStart) + xml.slice(cEnd);
  }
  const grid = findElement(xml, "w:tblGrid");
  if (grid) {
    const gridXml = xml.slice(grid.innerStart, grid.innerEnd);
    const cols = findElements(gridXml, "w:gridCol");
    if (idx < cols.length) {
      xml = xml.slice(0, grid.innerStart + cols[idx].outerStart) + xml.slice(grid.innerStart + cols[idx].outerEnd);
    }
  }
  return spliceTable(documentXml, table, xml);
}
