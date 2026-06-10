import { describe, it, expect } from "vitest";
import { runCss, paragraphCss, drawingCss } from "../src/cssMap";

describe("cssMap: run props → CSS", () => {
  it("maps weight, style, decoration, color, size, font, highlight, caps", () => {
    expect(runCss({
      bold: true, italic: false, underline: "single", strike: true,
      color: "FF0000", fontSize: 12, fonts: { ascii: "Arial" }, highlight: "yellow", caps: true,
    })).toEqual({
      fontWeight: "bold",
      fontStyle: "normal",
      textDecoration: "underline line-through",
      color: "#FF0000",
      fontSize: "12pt",
      fontFamily: "Arial, Helvetica, sans-serif",
      backgroundColor: "#ffff00",
      textTransform: "uppercase",
    });
  });

  it("ignores color=auto and renders superscript with a smaller size", () => {
    expect(runCss({ color: "auto" }).color).toBeUndefined();
    const sup = runCss({ vertAlign: "superscript" });
    expect(sup.verticalAlign).toBe("super");
    expect(sup.fontSize).toBe("0.8em");
  });

  it("run shading maps to background when there is no highlight", () => {
    expect(runCss({ shading: "FFFF00" }).backgroundColor).toBe("#FFFF00");
  });
});

describe("cssMap: paragraph props → CSS", () => {
  it("maps alignment, indent, spacing, auto line height", () => {
    expect(paragraphCss({
      alignment: "both",
      indent: { left: 720, firstLine: 360 },
      spacing: { before: 240, after: 120, line: 360, lineRule: "auto" },
    })).toEqual({
      textAlign: "justify",
      marginLeft: 48, // 720 twips @ 96dpi
      textIndent: 24,
      marginTop: 16,
      marginBottom: 8,
      lineHeight: 1.5, // 360/240
    });
  });

  it("hanging indent is a negative text-indent; exact line rule is px", () => {
    expect(paragraphCss({ indent: { hanging: 360 } }).textIndent).toBe(-24);
    expect(paragraphCss({ spacing: { line: 480, lineRule: "exact" } }).lineHeight).toBe("32px");
  });
});

describe("cssMap: drawing sizing", () => {
  it("EMU → px", () => {
    expect(drawingCss(914400, 457200)).toEqual({ width: 96, height: 48 });
    expect(drawingCss()).toEqual({});
  });
});

