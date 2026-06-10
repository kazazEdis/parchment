import { describe, it, expect } from "vitest";
import { parseStyles } from "../src/styles";
import { parseNumbering } from "../src/numbering";
import { parseDocument, type Paragraph } from "../src/model";
import { effectiveRunProps, effectiveParagraphProps, assignListNumbers } from "../src/resolve";

const STYLES = `<w:styles>
  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/><w:rFonts w:ascii="Calibri"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Strong"><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
  </w:style>
</w:styles>`;

const NUMBERING = `<w:numbering>
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8226;"/></w:lvl></w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

const sheet = parseStyles(STYLES);
const numbering = parseNumbering(NUMBERING);

describe("resolve: effective property cascade", () => {
  it("composes docDefaults → paragraph style → direct run", () => {
    const rPr = effectiveRunProps(sheet, numbering, { styleId: "Heading1" }, { italic: true });
    expect(rPr).toMatchObject({ bold: true, italic: true, fontSize: 16, fonts: { ascii: "Calibri" } });
  });

  it("a character style does NOT reset the paragraph style's size (the docDefaults-once fix)", () => {
    // Strong only sets color; size must remain 16 from Heading1, not fall back to docDefaults 11.
    const rPr = effectiveRunProps(sheet, numbering, { styleId: "Heading1" }, { styleId: "Strong" });
    expect(rPr.fontSize).toBe(16);
    expect(rPr.bold).toBe(true);
    expect(rPr.color).toBe("FF0000");
  });

  it("a table style's pPr applies to cell paragraphs (spacing-after reset, single line)", () => {
    // Outside a table: cell-less paragraph inherits docDefaults spacing-after 160 / line 259.
    const flow = effectiveParagraphProps(sheet, numbering, {});
    expect(flow.spacing?.after).toBe(160);
    expect(flow.spacing?.line).toBe(259);
    // Inside a TableGrid cell: the table style resets after→0 and line→240 (Word's cascade).
    const inCell = effectiveParagraphProps(sheet, numbering, {}, "TableGrid");
    expect(inCell.spacing?.after).toBe(0);
    expect(inCell.spacing?.line).toBe(240);
  });

  it("a paragraph's own pStyle still beats the table style", () => {
    // Heading1 sets after=0 before=240; the table style must not clobber the paragraph style.
    const eff = effectiveParagraphProps(sheet, numbering, { styleId: "Heading1" }, "TableGrid");
    expect(eff.spacing?.after).toBe(0);
    expect(eff.spacing?.before).toBe(240);
  });

  it("direct run formatting overrides the style", () => {
    const rPr = effectiveRunProps(sheet, numbering, { styleId: "Heading1" }, { bold: false, fontSize: 10 });
    expect(rPr.bold).toBe(false);
    expect(rPr.fontSize).toBe(10);
  });

  it("composes effective paragraph props (style + direct)", () => {
    const pPr = effectiveParagraphProps(sheet, numbering, { styleId: "Heading1", alignment: "center" });
    expect(pPr.keepNext).toBe(true);
    expect(pPr.alignment).toBe("center");
    // before/after from Heading1; line/lineRule inherited from docDefaults pPrDefault.
    expect(pPr.spacing).toEqual({ before: 240, after: 0, line: 259, lineRule: "auto" });
  });
});

describe("resolve: numbering counter pass", () => {
  const listPara = (numId: number, ilvl: number) =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`;
  const DOC =
    `<w:document><w:body>` +
    listPara(1, 0) + listPara(1, 0) + listPara(1, 1) + listPara(1, 1) +
    listPara(1, 0) + listPara(1, 1) + listPara(2, 0) +
    `</w:body></w:document>`;

  it("numbers a multi-level list with restart-on-deeper and bullets", () => {
    const model = parseDocument(DOC);
    const markers = assignListNumbers(model, numbering);
    const got = model.body.map((b) => markers.get(b as Paragraph));
    expect(got).toEqual(["1.", "2.", "2.a)", "2.b)", "3.", "3.a)", "•"]);
  });

  it("ignores paragraphs whose numId does not exist", () => {
    const model = parseDocument(`<w:document><w:body>${listPara(99, 0)}</w:body></w:document>`);
    const markers = assignListNumbers(model, numbering);
    expect(markers.size).toBe(0);
  });
});

