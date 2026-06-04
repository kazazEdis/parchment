import { describe, it, expect } from "vitest";
import { parseDocument, paragraphText, type Table, type Paragraph as P } from "../src/model";
import { insertRowAfter, deleteRow, appendColumn, deleteColumn } from "../src/tableEdit";

const DOC =
  `<w:document><w:body><w:tbl>` +
  `<w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>` +
  `<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>` +
  `<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>` +
  `</w:tbl></w:body></w:document>`;

const table = (xml: string): Table => parseDocument(xml).body[0] as Table;
const cellText = (t: Table, r: number, c: number): string => paragraphText(t.rows[r].cells[c].blocks[0] as P);

describe("tableEdit: rows", () => {
  it("inserts a blank row after the given row", () => {
    const t = table(insertRowAfter(DOC, table(DOC), 0));
    expect(t.rows).toHaveLength(3);
    expect(cellText(t, 0, 0)).toBe("A1"); // original
    expect(cellText(t, 1, 0)).toBe(""); // new blank row
    expect(cellText(t, 2, 0)).toBe("A2");
  });

  it("deletes a row (keeping at least one)", () => {
    const t = table(deleteRow(DOC, table(DOC), 0));
    expect(t.rows).toHaveLength(1);
    expect(cellText(t, 0, 0)).toBe("A2");
    // never deletes the last remaining row
    const one = table(deleteRow(`<w:document><w:body><w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>`, parseDocument(`<w:document><w:body><w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>`).body[0] as Table, 0));
    expect(one.rows).toHaveLength(1);
  });
});

describe("tableEdit: columns", () => {
  it("appends a blank column (gridCol + a cell per row)", () => {
    const t = table(appendColumn(DOC, table(DOC)));
    expect(t.grid).toHaveLength(3);
    expect(t.rows[0].cells).toHaveLength(3);
    expect(cellText(t, 0, 2)).toBe("");
  });

  it("deletes a column from the grid and every row", () => {
    const t = table(deleteColumn(DOC, table(DOC), 0));
    expect(t.grid).toHaveLength(1);
    expect(t.rows[0].cells).toHaveLength(1);
    expect(cellText(t, 0, 0)).toBe("B1");
    expect(cellText(t, 1, 0)).toBe("B2");
  });
});

