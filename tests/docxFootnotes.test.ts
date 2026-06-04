import { describe, it, expect } from "vitest";
import { parseFootnotes, footnoteRefs } from "../src/footnotes";
import { parseDocument, type Paragraph as P, type FootnoteRef } from "../src/model";
import { emitParagraph } from "../src/serialize";

const FOOTNOTES = `<w:footnotes>
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:r><w:t xml:space="preserve">See clause </w:t></w:r><w:r><w:t>4.2.</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>Net of VAT.</w:t></w:r></w:p></w:footnote>
</w:footnotes>`;

describe("footnotes: parse", () => {
  it("reads footnotes, skipping separators", () => {
    expect(parseFootnotes(FOOTNOTES)).toEqual([
      { id: "1", text: "See clause 4.2." },
      { id: "2", text: "Net of VAT." },
    ]);
    expect(parseFootnotes(undefined)).toEqual([]);
  });

  it("locates footnote references in the body", () => {
    const body = `<w:document><w:body><w:p><w:r><w:t>Total</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p></w:body></w:document>`;
    expect(footnoteRefs(body)).toEqual(["2"]);
  });
});

describe("footnotes: model marker round-trips", () => {
  it("captures w:footnoteReference as an inline and re-emits it", () => {
    const xml = `<w:document><w:body><w:p><w:r><w:t>Total</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p></w:body></w:document>`;
    const p = parseDocument(xml).body[0] as P;
    const ref = p.children.find((n): n is FootnoteRef => n.type === "footnoteRef");
    expect(ref).toEqual({ type: "footnoteRef", id: "2" });
    expect(emitParagraph(p)).toContain(`<w:footnoteReference w:id="2"/>`);
  });
});

