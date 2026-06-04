import { describe, it, expect } from "vitest";
import { Document, Packer, Paragraph, TextRun } from "docx";
import { parseDocument, paragraphText, type Paragraph as P } from "../src/model";
import { replaceInParagraph, markTracked } from "../src/edit";
import {
  fromPackage,
  replaceText,
  fillTemplate,
  getText,
  acceptAllChanges,
  rejectAllChanges,
  transformParagraphs,
  save,
} from "../src/doc";
import { readDocx } from "../src/opc";

describe("doc: split-run-aware replace (beats naive per-run)", () => {
  it("replaces a placeholder that Word split across runs", () => {
    const p = parseDocument(
      `<w:document><w:body><w:p>` +
        `<w:r><w:t xml:space="preserve">Hello {cust</w:t></w:r>` +
        `<w:r><w:rPr><w:b/></w:rPr><w:t>omer}!</w:t></w:r>` +
        `</w:p></w:body></w:document>`,
    ).body[0] as P;
    const { paragraph, count } = replaceInParagraph(p, "{customer}", "ACME");
    expect(count).toBe(1);
    expect(paragraphText(paragraph)).toBe("Hello ACME!");
  });

  it("falls back to per-run replace when non-run inlines are present", () => {
    const p = parseDocument(
      `<w:document><w:body><w:p><w:r><w:t xml:space="preserve">see {x}</w:t></w:r>` +
        `<w:hyperlink r:id="r1"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p></w:body></w:document>`,
    ).body[0] as P;
    const { paragraph, count } = replaceInParagraph(p, "{x}", "Y");
    expect(count).toBe(1);
    expect(paragraphText(paragraph)).toBe("see Ylink");
  });
});

async function fixture(children: Paragraph[]): Promise<Uint8Array> {
  return new Uint8Array(await Packer.toBuffer(new Document({ sections: [{ children }] })));
}

describe("doc: headless Document API on a real .docx", () => {
  it("fillTemplate replaces tokens everywhere and round-trips through save", async () => {
    const bytes = await fixture([
      new Paragraph({ children: [new TextRun("Offer {n} for {c}")] }),
      new Paragraph({ children: [new TextRun("Buyer: {c}")] }),
    ]);
    const doc = fromPackage(await readDocx(bytes));
    const { doc: filled, count } = fillTemplate(doc, { n: "2026-001", c: "ACME" });
    expect(count).toBe(3);
    expect(getText(filled)).toContain("Offer 2026-001 for ACME");
    expect(getText(filled)).toContain("Buyer: ACME");

    const reopened = fromPackage(await readDocx(await save(filled)));
    expect(getText(reopened)).toContain("Offer 2026-001 for ACME");
  });

  it("replaceText is a no-op (same Doc, count 0) when nothing matches", async () => {
    const doc = fromPackage(await readDocx(await fixture([new Paragraph({ children: [new TextRun("nothing")] })])));
    const { doc: d2, count } = replaceText(doc, "{missing}", "x");
    expect(count).toBe(0);
    expect(d2).toBe(doc);
  });

  it("accept / reject all tracked changes across the whole document", async () => {
    let doc = fromPackage(await readDocx(await fixture([new Paragraph({ children: [new TextRun("kept")] })])));
    doc = transformParagraphs(doc, () => true, (p) => markTracked(p, { type: "ins" as const, author: "A" }));

    expect(getText(rejectAllChanges(doc))).toBe(""); // insertion rejected → text gone
    expect(getText(acceptAllChanges(doc))).toBe("kept"); // insertion accepted → plain text
  });
});

