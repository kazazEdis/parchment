import { describe, it, expect } from "vitest";
import { parseDocument, type Paragraph, type Drawing } from "../src/model";

const doc = (body: string) => `<w:document xmlns:w="x" xmlns:wp="y" xmlns:a="z" xmlns:pic="p" xmlns:r="rr"><w:body>${body}</w:body></w:document>`;

const drawing = (srcRect: string) => doc(`<w:p><w:r><w:drawing><wp:inline>
  <wp:extent cx="600000" cy="300000"/><wp:docPr id="1" name="logo"/>
  <a:graphic><a:graphicData><pic:pic><pic:blipFill>
    <a:blip r:embed="rId3"/>${srcRect}<a:stretch><a:fillRect/></a:stretch>
  </pic:blipFill></pic:pic></a:graphicData></a:graphic>
</wp:inline></w:drawing></w:r></w:p>`);

const firstDrawing = (xml: string): Drawing =>
  (parseDocument(xml).body[0] as Paragraph).children.find((n): n is Drawing => n.type === "drawing")!;

describe("drawing: a:srcRect crop", () => {
  it("parses srcRect into edge-crop fractions (ST_Percentage, 100000 = 100%)", () => {
    const d = firstDrawing(drawing(`<a:srcRect l="923" t="11468" r="85300" b="79118"/>`));
    expect(d.crop).toBeDefined();
    expect(d.crop!.l).toBeCloseTo(0.00923, 5);
    expect(d.crop!.t).toBeCloseTo(0.11468, 5);
    expect(d.crop!.r).toBeCloseTo(0.853, 5);
    expect(d.crop!.b).toBeCloseTo(0.79118, 5);
  });

  it("no crop when srcRect is absent", () => {
    expect(firstDrawing(drawing(``)).crop).toBeUndefined();
  });

  it("no crop when all edges are zero / empty", () => {
    expect(firstDrawing(drawing(`<a:srcRect/>`)).crop).toBeUndefined();
    expect(firstDrawing(drawing(`<a:srcRect l="0" t="0" r="0" b="0"/>`)).crop).toBeUndefined();
  });
});
