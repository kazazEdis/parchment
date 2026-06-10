// Theme font resolution: w:asciiTheme tokens (minorHAnsi/majorHAnsi) resolve to the theme part's
// major/minor latin typefaces in the run cascade, while staying tokens on round-trip (serialize).
import { describe, it, expect } from "vitest";
import { parseStyles, parseThemeFonts, resolveThemeFont } from "../src/styles";
import { effectiveRunProps } from "../src/resolve";
import { parseRunProps } from "../src/props";
import { emitRunProps } from "../src/serialize";
import { fontFamilyCss } from "../src/cssMap";

const THEME = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements><a:fontScheme name="Office">
    <a:majorFont><a:latin typeface="Cambria"/></a:majorFont>
    <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
  </a:fontScheme></a:themeElements></a:theme>`;

const STYLES = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>`;

describe("theme fonts", () => {
  it("parses major/minor latin typefaces from the theme part", () => {
    expect(parseThemeFonts(THEME)).toEqual({ major: "Cambria", minor: "Calibri" });
    expect(parseThemeFonts(undefined)).toBeUndefined();
  });

  it("maps tokens: minor* → minor, major* → major", () => {
    const tf = { major: "Cambria", minor: "Calibri" };
    expect(resolveThemeFont(tf, "minorHAnsi")).toBe("Calibri");
    expect(resolveThemeFont(tf, "majorHAnsi")).toBe("Cambria");
    expect(resolveThemeFont(tf, "minorAscii")).toBe("Calibri");
    expect(resolveThemeFont(tf, undefined)).toBeUndefined();
    expect(resolveThemeFont(undefined, "minorHAnsi")).toBeUndefined();
  });

  it("docDefaults asciiTheme resolves to a concrete ascii in effectiveRunProps", () => {
    const sheet = parseStyles(STYLES, THEME);
    const eff = effectiveRunProps(sheet, undefined, {}, {});
    expect(eff.fonts?.ascii).toBe("Calibri");
    expect(eff.fonts?.asciiTheme).toBe("minorHAnsi"); // token preserved alongside
  });

  it("a direct concrete w:ascii beats the inherited theme token", () => {
    const sheet = parseStyles(STYLES, THEME);
    const direct = parseRunProps(`<w:rFonts w:ascii="Arial"/>`);
    const eff = effectiveRunProps(sheet, undefined, {}, direct);
    expect(eff.fonts?.ascii).toBe("Arial");
  });

  it("without a theme part the token stays unresolved (no font emitted)", () => {
    const sheet = parseStyles(STYLES); // no theme XML
    const eff = effectiveRunProps(sheet, undefined, {}, {});
    expect(eff.fonts?.ascii).toBeUndefined();
  });

  it("round-trips asciiTheme through parse → serialize without baking a concrete font", () => {
    const p = parseRunProps(`<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>`);
    const xml = emitRunProps(p);
    expect(xml).toContain(`w:asciiTheme="minorHAnsi"`);
    expect(xml).toContain(`w:hAnsiTheme="minorHAnsi"`);
    expect(xml).not.toContain(`w:ascii="`);
  });

  it("fontFamilyCss appends metric-compatible + generic fallbacks", () => {
    expect(fontFamilyCss("Calibri")).toBe("Calibri, Carlito, sans-serif");
    expect(fontFamilyCss("Times New Roman")).toBe(`"Times New Roman", Liberation Serif, Georgia, serif`);
    expect(fontFamilyCss("SomeUnknownFont")).toBe("SomeUnknownFont, sans-serif");
  });
});
