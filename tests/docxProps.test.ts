import { describe, it, expect } from "vitest";
import {
  parseRunProps,
  parseParagraphProps,
  mergeRunProps,
  mergeParagraphProps,
} from "../src/props";

describe("props: run properties", () => {
  it("parses common rPr toggles, size, color, underline, fonts", () => {
    const rPr = parseRunProps(
      `<w:b/><w:i w:val="false"/><w:sz w:val="24"/><w:color w:val="FF0000"/><w:u w:val="single"/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>`,
    );
    expect(rPr).toEqual({
      bold: true,
      italic: false, // explicit off
      fontSize: 12, // 24 half-points
      color: "FF0000",
      underline: "single",
      fonts: { ascii: "Calibri", hAnsi: "Calibri" },
    });
  });

  it("treats a bare toggle as on and w:u without val as single", () => {
    expect(parseRunProps(`<w:b/>`).bold).toBe(true);
    expect(parseRunProps(`<w:u/>`).underline).toBe("single");
    expect(parseRunProps(`<w:b w:val="0"/>`).bold).toBe(false);
  });
});

describe("props: paragraph properties", () => {
  it("parses style ref, alignment, indent, spacing, numbering", () => {
    const pPr = parseParagraphProps(
      `<w:pStyle w:val="Heading1"/><w:jc w:val="center"/><w:ind w:left="720" w:hanging="360"/>` +
      `<w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/>` +
      `<w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr>`,
    );
    expect(pPr.styleId).toBe("Heading1");
    expect(pPr.alignment).toBe("center");
    expect(pPr.indent).toEqual({ left: 720, hanging: 360 });
    expect(pPr.spacing).toEqual({ before: 120, after: 240, line: 360, lineRule: "auto" });
    expect(pPr.numbering).toEqual({ numId: 3, level: 1 });
  });

  it("captures the paragraph-mark run props (w:pPr/w:rPr)", () => {
    const pPr = parseParagraphProps(`<w:jc w:val="left"/><w:rPr><w:sz w:val="20"/></w:rPr>`);
    expect(pPr.markRunProps?.fontSize).toBe(10);
  });
});

describe("props: cascade merge (override wins, nested shallow-merge)", () => {
  it("merges run props, later overriding", () => {
    expect(mergeRunProps({ bold: true, fonts: { ascii: "A" } }, { italic: true, fonts: { hAnsi: "B" } }))
      .toEqual({ bold: true, italic: true, fonts: { ascii: "A", hAnsi: "B" } });
    expect(mergeRunProps({ bold: true }, { bold: false }).bold).toBe(false);
  });

  it("merges paragraph props, nested indent/spacing shallow-merge", () => {
    const out = mergeParagraphProps(
      { spacing: { after: 160 }, indent: { left: 100 } },
      { spacing: { before: 240, after: 0 } },
    );
    expect(out.spacing).toEqual({ before: 240, after: 0 });
    expect(out.indent).toEqual({ left: 100 });
  });
});

