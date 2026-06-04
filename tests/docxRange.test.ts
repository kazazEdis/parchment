import { describe, it, expect } from "vitest";
import { parseDocument, paragraphText, type Paragraph as P, type Run } from "../src/model";
import { formatRange, spliceRunRange, rangeUniform, rangeUnderlined, paragraphLength } from "../src/edit";

const runsOf = (p: P): Run[] => p.children.filter((n): n is Run => n.type === "run");
const para = (xml: string): P => parseDocument(`<w:document><w:body><w:p>${xml}</w:p></w:body></w:document>`).body[0] as P;

describe("range: formatRange splits runs at the selection boundary", () => {
  it("splits a single run into before / formatted / after", () => {
    const p = formatRange(para(`<w:r><w:t>ABCDE</w:t></w:r>`), 1, 3, { bold: true });
    expect(runsOf(p).map((r) => r.text)).toEqual(["A", "BC", "DE"]);
    expect(runsOf(p).map((r) => r.rPr.bold)).toEqual([undefined, true, undefined]);
    expect(paragraphText(p)).toBe("ABCDE");
  });

  it("applies across run boundaries, preserving each run's own formatting", () => {
    const p = formatRange(para(`<w:r><w:t>AB</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>CD</w:t></w:r>`), 0, 3, { bold: true });
    const rs = runsOf(p);
    expect(rs.map((r) => r.text)).toEqual(["AB", "C", "D"]);
    expect(rs[0].rPr).toMatchObject({ bold: true });
    expect(rs[0].rPr.italic).toBeUndefined();
    expect(rs[1].rPr).toMatchObject({ bold: true, italic: true });
    expect(rs[2].rPr).toMatchObject({ italic: true });
    expect(rs[2].rPr.bold).toBeUndefined();
  });

  it("is a no-op on an empty range", () => {
    const p = para(`<w:r><w:t>ABCDE</w:t></w:r>`);
    expect(paragraphText(formatRange(p, 2, 2, { bold: true }))).toBe("ABCDE");
  });
});

describe("range: spliceRunRange inserts / deletes / replaces", () => {
  const p = para(`<w:r><w:t>ABCDE</w:t></w:r>`);

  it("deletes a range", () => {
    expect(paragraphText(spliceRunRange(p, 1, 3, []))).toBe("ADE");
  });

  it("inserts runs at a caret, keeping their formatting", () => {
    const out = spliceRunRange(p, 2, 2, [{ type: "run", rPr: { bold: true }, text: "XX" }]);
    expect(paragraphText(out)).toBe("ABXXCDE");
    expect(runsOf(out).find((r) => r.text === "XX")!.rPr.bold).toBe(true);
  });

  it("replaces a range", () => {
    expect(paragraphText(spliceRunRange(p, 0, 2, [{ type: "run", rPr: {}, text: "Z" }]))).toBe("ZCDE");
  });
});

describe("range: selection-state queries", () => {
  const p = para(`<w:r><w:t>A</w:t></w:r><w:r><w:rPr><w:b/><w:u w:val="single"/></w:rPr><w:t>BC</w:t></w:r>`);

  it("rangeUniform reports whether the whole range has a toggle", () => {
    expect(rangeUniform(p, 1, 3, "bold")).toBe(true);
    expect(rangeUniform(p, 0, 3, "bold")).toBe(false);
  });

  it("rangeUnderlined detects a fully-underlined range", () => {
    expect(rangeUnderlined(p, 1, 3)).toBe(true);
    expect(rangeUnderlined(p, 0, 3)).toBe(false);
  });

  it("paragraphLength counts run characters", () => {
    expect(paragraphLength(p)).toBe(3);
  });
});

