import { describe, it, expect } from "vitest";
import { parseDocument, paragraphText, type Paragraph as P, type Run } from "../src/model";
import { splitParagraph, mergeParagraphs, setListLevel } from "../src/edit";

const runsOf = (p: P): Run[] => p.children.filter((n): n is Run => n.type === "run");
const para = (xml: string): P => parseDocument(`<w:document><w:body><w:p>${xml}</w:p></w:body></w:document>`).body[0] as P;

describe("block: splitParagraph", () => {
  it("splits a single run at the offset, keeping pPr on both halves", () => {
    const p = para(`<w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>ABCDE</w:t></w:r>`);
    const [a, b] = splitParagraph(p, 2);
    expect(paragraphText(a)).toBe("AB");
    expect(paragraphText(b)).toBe("CDE");
    expect(a.pPr.styleId).toBe("Body");
    expect(b.pPr.styleId).toBe("Body");
  });

  it("splits across runs, preserving formatting on each side", () => {
    const p = para(`<w:r><w:t>AB</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>CD</w:t></w:r>`);
    const [a, b] = splitParagraph(p, 3);
    expect(paragraphText(a)).toBe("ABC");
    expect(paragraphText(b)).toBe("D");
    expect(runsOf(a).find((r) => r.text === "C")!.rPr.bold).toBe(true);
    expect(runsOf(b)[0].rPr.bold).toBe(true);
  });
});

describe("block: mergeParagraphs", () => {
  it("concatenates b's runs onto a, keeping a's pPr", () => {
    const a = para(`<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>AB</w:t></w:r>`);
    const b = para(`<w:r><w:t>CD</w:t></w:r>`);
    const merged = mergeParagraphs(a, b);
    expect(paragraphText(merged)).toBe("ABCD");
    expect(merged.pPr.alignment).toBe("center");
  });
});

describe("block: setListLevel (indent / outdent)", () => {
  it("changes a list paragraph's level, clamped at 0", () => {
    const list = para(`<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r>`);
    expect(setListLevel(list, 1).pPr.numbering!.level).toBe(1);
    expect(setListLevel(list, -1).pPr.numbering!.level).toBe(0);
  });

  it("adjusts left indent for a non-list paragraph", () => {
    const p = para(`<w:r><w:t>x</w:t></w:r>`);
    expect(setListLevel(p, 1).pPr.indent!.left).toBe(720);
    expect(setListLevel(setListLevel(p, 1), -1).pPr.indent!.left).toBe(0);
  });
});

