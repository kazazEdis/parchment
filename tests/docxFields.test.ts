import { describe, it, expect } from "vitest";
import { parseFieldInstruction, generateToc, computePageField } from "../src/fields";
import { parseDocument, paragraphText, type Paragraph as P } from "../src/model";

describe("fields: parse instruction", () => {
  it("parses type, args, and switches", () => {
    expect(parseFieldInstruction("PAGE")).toEqual({ type: "PAGE", args: [], switches: {} });
    expect(parseFieldInstruction('TOC \\o "1-3" \\h')).toEqual({ type: "TOC", args: [], switches: { "\\o": "1-3", "\\h": true } });
    expect(parseFieldInstruction("REF _Ref1 \\h")).toEqual({ type: "REF", args: ["_Ref1"], switches: { "\\h": true } });
    expect(parseFieldInstruction('HYPERLINK "http://x"')).toEqual({ type: "HYPERLINK", args: ["http://x"], switches: {} });
  });
});

describe("fields: TOC + page values", () => {
  it("generates TOC entries from heading styles", () => {
    const xml =
      `<w:document><w:body>` +
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Intro</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Details</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>body text</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    expect(generateToc(parseDocument(xml))).toEqual([
      { level: 1, text: "Intro", styleId: "Heading1" },
      { level: 2, text: "Details", styleId: "Heading2" },
    ]);
  });

  it("computePageField resolves PAGE / NUMPAGES", () => {
    expect(computePageField("PAGE", 3, 7)).toBe("3");
    expect(computePageField("NUMPAGES", 3, 7)).toBe("7");
  });
});

describe("fields: fldSimple cached result renders", () => {
  it("unwraps a simple field so its result text shows", () => {
    const xml = `<w:document><w:body><w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p></w:body></w:document>`;
    expect(paragraphText(parseDocument(xml).body[0] as P)).toBe("Page 3");
  });
});

