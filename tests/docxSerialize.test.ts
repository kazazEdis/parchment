import { describe, it, expect } from "vitest";
import { Document, Packer, Paragraph as DocxParagraph, TextRun } from "docx";
import {
  emitParagraph,
  emitRun,
  emitInline,
  patchParagraph,
  patchAll,
  setParagraphText,
} from "../src/serialize";
import { parseDocument, paragraphText, type Paragraph as P, type Run } from "../src/model";
import { readDocx, writeDocx, getPartText, setPartText } from "../src/opc";

const DOC =
  `<w:document><w:body>` +
  `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:before="240" w:after="0"/><w:jc w:val="center"/></w:pPr>` +
  `<w:r><w:rPr><w:b/><w:i w:val="false"/><w:sz w:val="28"/><w:color w:val="FF0000"/><w:u w:val="single"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>Heading</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t xml:space="preserve">plain text</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>third</w:t></w:r></w:p>` +
  `</w:body></w:document>`;

describe("serialize: emit ∘ parse is identity on modelled props", () => {
  const model = parseDocument(DOC);

  it("round-trips a rich paragraph (style, spacing, alignment, run formatting)", () => {
    const p0 = model.body[0] as P;
    const reparsed = parseDocument(`<w:document><w:body>${emitParagraph(p0)}</w:body></w:document>`).body[0] as P;
    expect(reparsed.pPr).toEqual(p0.pPr);
    expect(reparsed.children).toEqual(p0.children);
  });

  it("round-trips tabs and line breaks inside a run", () => {
    const run: Run = { type: "run", rPr: {}, text: "a\tb\nc" };
    const xml = emitRun(run);
    expect(xml).toContain("<w:tab/>");
    expect(xml).toContain("<w:br/>");
    const reparsed = parseDocument(`<w:document><w:body><w:p>${xml}</w:p></w:body></w:document>`).body[0] as P;
    expect((reparsed.children[0] as Run).text).toBe("a\tb\nc");
  });

  it("drops drawings on re-emit (documented v1 limitation)", () => {
    expect(emitInline({ type: "drawing", anchored: false })).toBe("");
  });
});

describe("serialize: span-splice patching (preserve-and-patch)", () => {
  const model = parseDocument(DOC);

  it("patches one paragraph and leaves all bytes outside its span untouched", () => {
    const third = model.body[2] as P;
    const out = patchParagraph(DOC, setParagraphText(third, "THIRD"));
    expect(out.slice(0, third.source.start)).toBe(DOC.slice(0, third.source.start)); // prefix intact
    const reparsed = parseDocument(out);
    expect(reparsed.body.map((b) => paragraphText(b as P))).toEqual(["Heading", "plain text", "THIRD"]);
  });

  it("patchAll applies multiple edits with stable offsets", () => {
    const p0 = model.body[0] as P;
    const p2 = model.body[2] as P;
    const out = patchAll(DOC, [
      { span: p0.source, xml: emitParagraph(setParagraphText(p0, "H2")) },
      { span: p2.source, xml: emitParagraph(setParagraphText(p2, "T3")) },
    ]);
    expect(parseDocument(out).body.map((b) => paragraphText(b as P))).toEqual(["H2", "plain text", "T3"]);
  });

  it("setParagraphText keeps the first run's formatting", () => {
    const p0 = model.body[0] as P;
    const edited = setParagraphText(p0, "X");
    expect((edited.children[0] as Run).rPr.bold).toBe(true);
    expect((edited.children[0] as Run).rPr.color).toBe("FF0000");
  });
});

describe("serialize: full-package round-trip", () => {
  async function fixture(): Promise<Uint8Array> {
    const doc = new Document({
      sections: [{ children: [
        new DocxParagraph({ children: [new TextRun("Number: {n}")] }),
        new DocxParagraph({ children: [new TextRun("Customer: {c}")] }),
      ] }],
    });
    return new Uint8Array(await Packer.toBuffer(doc));
  }

  it("editing one paragraph rewrites only document.xml; other parts stay byte-identical", async () => {
    const pkg = await readDocx(await fixture());
    const xml0 = getPartText(pkg, "word/document.xml")!;
    const model = parseDocument(xml0);

    const target = model.body.find((b): b is P => b.type === "paragraph" && paragraphText(b).includes("Customer"))!;
    setPartText(pkg, "word/document.xml", patchParagraph(xml0, setParagraphText(target, "Customer: ACME")));

    const otherParts = new Map(
      pkg.order.filter((p) => p !== "word/document.xml").map((p) => [p, Array.from(pkg.parts.get(p)!.bytes)] as const),
    );

    const reopened = await readDocx(await writeDocx(pkg));
    const reModel = parseDocument(getPartText(reopened, "word/document.xml")!);
    expect(reModel.body.map((b) => paragraphText(b as P))).toEqual(["Number: {n}", "Customer: ACME"]);
    for (const [path, bytes] of otherParts) {
      expect(Array.from(reopened.parts.get(path)!.bytes), `part drifted: ${path}`).toEqual(bytes);
    }
  });
});

