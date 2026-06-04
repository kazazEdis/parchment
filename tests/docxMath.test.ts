import { describe, it, expect } from "vitest";
import { ommlToMathML } from "../src/math";
import { parseDocument, type Paragraph as P, type MathInline } from "../src/model";
import { emitParagraph } from "../src/serialize";

describe("math: OMML → MathML", () => {
  it("converts a fraction", () => {
    const omml = `<m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>`;
    const ml = ommlToMathML(omml);
    expect(ml).toContain("<math");
    expect(ml).toContain("<mfrac>");
    expect(ml).toContain("<mi>a</mi>");
    expect(ml).toContain("<mi>b</mi>");
  });

  it("converts a superscript with a numeric exponent", () => {
    const omml = `<m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>`;
    const ml = ommlToMathML(omml);
    expect(ml).toContain("<msup>");
    expect(ml).toContain("<mi>x</mi>");
    expect(ml).toContain("<mn>2</mn>");
  });
});

describe("math: model captures + serialize preserves oMath", () => {
  it("keeps the original OMML verbatim", () => {
    const xml = `<w:document><w:body><w:p><w:r><w:t>E=</w:t></w:r><m:oMath><m:r><m:t>mc</m:t></m:r></m:oMath></w:p></w:body></w:document>`;
    const p = parseDocument(xml).body[0] as P;
    const math = p.children.find((n): n is MathInline => n.type === "math");
    expect(math?.omml).toContain("m:oMath");
    expect(emitParagraph(p)).toContain("<m:oMath>");
  });
});

