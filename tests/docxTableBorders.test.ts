import { describe, it, expect } from "vitest";
import { parseDocument, type Table } from "../src/model";
import { parseStyles, resolveTableStyleBorders } from "../src/styles";

const doc = (body: string) => `<w:document xmlns:w="x"><w:body>${body}</w:body></w:document>`;

describe("table borders: parsing", () => {
  it("parses inline w:tblBorders, w:tblStyle, cell w:tcBorders and w:shd", () => {
    const xml = doc(`
      <w:tbl>
        <w:tblPr>
          <w:tblStyle w:val="TableGrid"/>
          <w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="none" w:sz="0" w:color="auto"/></w:tblBorders>
        </w:tblPr>
        <w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>
        <w:tr><w:tc>
          <w:tcPr><w:tcBorders><w:bottom w:val="single" w:sz="8" w:color="FF0000"/></w:tcBorders><w:shd w:fill="C0FFEE"/></w:tcPr>
          <w:p><w:r><w:t>x</w:t></w:r></w:p>
        </w:tc></w:tr>
      </w:tbl>`);
    const t = parseDocument(xml).body[0] as Table;
    expect(t.type).toBe("table");
    expect(t.styleId).toBe("TableGrid");
    expect(t.borders?.top).toMatchObject({ val: "single", sz: 4 });
    expect(t.borders?.insideV?.val).toBe("none");
    const cell = t.rows[0].cells[0];
    expect(cell.props.borders?.bottom).toMatchObject({ val: "single", sz: 8, color: "FF0000" });
    expect(cell.props.shd).toBe("C0FFEE");
  });

  it("ignores an auto/empty shd fill", () => {
    const xml = doc(`<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:shd w:fill="auto"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`);
    const t = parseDocument(xml).body[0] as Table;
    expect(t.rows[0].cells[0].props.shd).toBeUndefined();
  });

  it("a table with no border info parses to undefined borders (renderer draws nothing)", () => {
    const xml = doc(`<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>
      <w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`);
    const t = parseDocument(xml).body[0] as Table;
    expect(t.borders).toBeUndefined();
    expect(t.styleId).toBeUndefined();
  });
});

describe("table borders: style resolution", () => {
  it("resolveTableStyleBorders reads a table style's w:tblBorders (e.g. TableGrid)", () => {
    const styles = `<w:styles xmlns:w="x"><w:style w:type="table" w:styleId="TableGrid">
      <w:tblPr><w:tblBorders>
        <w:top w:val="single" w:sz="4" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:color="auto"/>
      </w:tblBorders></w:tblPr></w:style></w:styles>`;
    const b = resolveTableStyleBorders(parseStyles(styles), "TableGrid");
    expect(b?.top?.val).toBe("single");
    expect(b?.insideH).toMatchObject({ val: "single", sz: 4 });
  });

  it("returns undefined for an unknown style", () => {
    expect(resolveTableStyleBorders(parseStyles(`<w:styles xmlns:w="x"/>`), "Nope")).toBeUndefined();
  });
});
