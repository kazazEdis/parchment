import { describe, it, expect } from "vitest";
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from "docx";
import { parseDocument, paragraphText, documentText, type Paragraph as P, type Table } from "../src/model";
import { readDocx, getPartText } from "../src/opc";
import { replaceSpan, findElement } from "../src/xml";

const DOC =
  `<w:document><w:body>` +
  `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Title</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:hyperlink r:id="rId5"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>` +
  `<w:tbl>` +
  `<w:tblGrid><w:gridCol w:w="4675"/><w:gridCol w:w="4675"/></w:tblGrid>` +
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>` +
  `<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>` +
  `</w:tbl>` +
  `<w:p><w:r><w:t>End</w:t></w:r></w:p>` +
  `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"/>` +
  `<w:headerReference w:type="default" r:id="rId7"/></w:sectPr>` +
  `</w:body></w:document>`;

describe("model: parse document.xml structure", () => {
  const model = parseDocument(DOC);

  it("returns blocks (paragraphs + table), excluding sectPr", () => {
    expect(model.body.map((b) => b.type)).toEqual(["paragraph", "paragraph", "table", "paragraph"]);
  });

  it("parses paragraph style + run formatting", () => {
    const p0 = model.body[0] as P;
    expect(p0.pPr.styleId).toBe("Heading1");
    expect(p0.children).toHaveLength(1);
    expect(p0.children[0]).toMatchObject({ type: "run", text: "Title", rPr: { bold: true } });
  });

  it("preserves significant whitespace and parses hyperlinks", () => {
    const p1 = model.body[1] as P;
    expect((p1.children[0] as { text: string }).text).toBe("Hello "); // xml:space=preserve trailing space
    const link = p1.children[1];
    expect(link).toMatchObject({ type: "hyperlink", rId: "rId5" });
    expect(paragraphText(p1)).toBe("Hello link");
  });

  it("parses the table grid + cells (recursively)", () => {
    const tbl = model.body[2] as Table;
    expect(tbl.grid).toEqual([4675, 4675]);
    expect(tbl.rows).toHaveLength(1);
    expect(tbl.rows[0].cells).toHaveLength(2);
    expect(tbl.rows[0].cells[0].props.width).toEqual({ value: 4675, type: "dxa" });
    expect(paragraphText(tbl.rows[0].cells[0].blocks[0] as P)).toBe("A1");
  });

  it("parses section properties", () => {
    expect(model.section?.pageSize).toEqual({ width: 11906, height: 16838, orient: undefined });
    expect(model.section?.margins?.top).toBe(1417);
    expect(model.section?.headerRefs).toEqual([{ type: "default", rId: "rId7" }]);
  });

  it("documentText flattens body + table text in order", () => {
    expect(documentText(model)).toBe("Title\nHello link\nA1\nB1\nEnd");
  });
});

describe("model: source spans drive preserve-and-patch", () => {
  const model = parseDocument(DOC);

  it("each block's source span reproduces its exact XML (identity patch)", () => {
    for (const b of model.body) {
      // Replacing a span with the very bytes it points at must be a no-op.
      const slice = DOC.slice(b.source.start, b.source.end);
      const span = findElement(DOC, b.type === "paragraph" ? "w:p" : "w:tbl", { from: b.source.start })!;
      expect(span.outerStart).toBe(b.source.start);
      expect(replaceSpan(DOC, { ...span, outerStart: b.source.start, outerEnd: b.source.end }, slice)).toBe(DOC);
    }
  });

  it("patching one paragraph by its span leaves the rest intact", () => {
    const end = model.body[3] as P;
    const patched = DOC.slice(0, end.source.start) + `<w:p><w:r><w:t>Fin</w:t></w:r></w:p>` + DOC.slice(end.source.end);
    const reparsed = parseDocument(patched);
    expect(documentText(reparsed)).toBe("Title\nHello link\nA1\nB1\nFin");
    expect((reparsed.body[0] as P).children[0]).toMatchObject({ text: "Title" }); // others untouched
  });
});

describe("model: parses real docx-lib output", () => {
  it("reads paragraphs + run formatting from a generated .docx", async () => {
    const doc = new Document({
      sections: [{ children: [
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Report")] }),
        new Paragraph({ children: [new TextRun({ text: "Bold", bold: true }), new TextRun(" plain")] }),
      ] }],
    });
    const pkg = await readDocx(new Uint8Array(await Packer.toBuffer(doc)));
    const xml = getPartText(pkg, "word/document.xml")!;
    const model = parseDocument(xml);

    const text = documentText(model);
    expect(text).toContain("Report");
    expect(text).toContain("Bold plain");
    // The bold run is modelled with rPr.bold.
    const all = model.body.filter((b): b is P => b.type === "paragraph").flatMap((p) => p.children);
    expect(all.some((n) => n.type === "run" && n.text === "Bold" && n.rPr.bold === true)).toBe(true);
  });
});

