import { describe, it, expect } from "vitest";
import { parseStyles, resolveParagraphStyle, resolveRunStyle } from "../src/styles";

const STYLES = `<w:styles>
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:sz w:val="22"/><w:rFonts w:ascii="Calibri"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="X">
    <w:pPr><w:rPr><w:i/></w:rPr></w:pPr><w:rPr><w:b/></w:rPr>
  </w:style>
</w:styles>`;

describe("styles: parse styles.xml", () => {
  const sheet = parseStyles(STYLES);

  it("reads docDefaults", () => {
    expect(sheet.docDefaults.rPr.fontSize).toBe(11); // sz 22 → 11pt
    expect(sheet.docDefaults.rPr.fonts?.ascii).toBe("Calibri");
    expect(sheet.docDefaults.pPr.spacing?.after).toBe(160);
  });

  it("reads style defs + default paragraph style", () => {
    expect([...sheet.styles.keys()].sort()).toEqual(["Heading1", "Normal", "Strong", "X"]);
    expect(sheet.styles.get("Heading1")!.name).toBe("heading 1");
    expect(sheet.styles.get("Heading1")!.basedOn).toBe("Normal");
    expect(sheet.defaultParagraphStyleId).toBe("Normal");
  });

  it("does not confuse the paragraph-mark rPr with the style-level rPr", () => {
    const x = sheet.styles.get("X")!;
    expect(x.rPr.bold).toBe(true); // style-level rPr (after pPr)
    expect(x.rPr.italic).toBeUndefined();
    expect(x.pPr.markRunProps?.italic).toBe(true); // pPr/rPr mark props
  });
});

describe("styles: cascade resolution", () => {
  const sheet = parseStyles(STYLES);

  it("resolves a paragraph style through docDefaults + basedOn chain", () => {
    const { pPr, rPr } = resolveParagraphStyle(sheet, "Heading1");
    expect(rPr.bold).toBe(true);
    expect(rPr.fontSize).toBe(16); // Heading1 sz 32 overrides docDefaults
    expect(rPr.fonts?.ascii).toBe("Calibri"); // inherited from docDefaults
    expect(pPr.keepNext).toBe(true);
    expect(pPr.spacing).toEqual({ before: 240, after: 0 }); // after overrides docDefaults 160
  });

  it("resolves a character style on top of docDefaults", () => {
    const rPr = resolveRunStyle(sheet, "Strong");
    expect(rPr.bold).toBe(true);
    expect(rPr.fontSize).toBe(11);
    expect(rPr.fonts?.ascii).toBe("Calibri");
  });

  it("unknown style id resolves to docDefaults only", () => {
    expect(resolveRunStyle(sheet, "Nope").fontSize).toBe(11);
    expect(resolveRunStyle(sheet, undefined).fontSize).toBe(11);
  });
});

