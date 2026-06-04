import { describe, it, expect } from "vitest";
import { parseNumbering, getLevel, formatMarker, toRoman, toLetter } from "../src/numbering";

const NUMBERING = `<w:numbering>
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8226;"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>
</w:numbering>`;

describe("numbering: parse + resolve levels", () => {
  const num = parseNumbering(NUMBERING);

  it("indexes abstractNums and nums", () => {
    expect([...num.abstractNums.keys()].sort()).toEqual([0, 1]);
    expect([...num.nums.keys()].sort()).toEqual([1, 2, 3]);
    expect(num.abstractNums.get(0)!.size).toBe(2);
  });

  it("resolves a level through num → abstractNum", () => {
    expect(getLevel(num, 1, 0)).toMatchObject({ numFmt: "decimal", lvlText: "%1.", start: 1 });
    expect(getLevel(num, 1, 1)!.numFmt).toBe("lowerLetter");
    expect(getLevel(num, 2, 0)!.isBullet).toBe(true);
  });

  it("applies a startOverride", () => {
    expect(getLevel(num, 3, 0)!.start).toBe(5);
  });

  it("returns undefined for an unknown numId/level", () => {
    expect(getLevel(num, 99, 0)).toBeUndefined();
    expect(getLevel(num, 1, 5)).toBeUndefined();
  });
});

describe("numbering: marker formatting", () => {
  const num = parseNumbering(NUMBERING);

  it("formats decimal and multi-level markers", () => {
    expect(formatMarker(num, 1, 0, [1])).toBe("1.");
    expect(formatMarker(num, 1, 0, [3])).toBe("3.");
    expect(formatMarker(num, 1, 1, [2, 3])).toBe("2.c)"); // %1=2 decimal, %2=3 lowerLetter
  });

  it("renders a bullet as its literal glyph", () => {
    expect(formatMarker(num, 2, 0, [1])).toBe("•");
  });
});

describe("numbering: integer → glyph", () => {
  it("roman numerals", () => {
    expect(toRoman(1)).toBe("I");
    expect(toRoman(4)).toBe("IV");
    expect(toRoman(9)).toBe("IX");
    expect(toRoman(2026)).toBe("MMXXVI");
  });
  it("spreadsheet-style letters (bijective base-26)", () => {
    expect(toLetter(1)).toBe("A");
    expect(toLetter(26)).toBe("Z");
    expect(toLetter(27)).toBe("AA");
    expect(toLetter(52)).toBe("AZ");
  });
});

