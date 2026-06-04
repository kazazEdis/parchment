import { describe, it, expect } from "vitest";
import { headerXml, footerXml } from "../src/headerFooter";
import { parseDocument, parseContainer, paragraphText, type Paragraph as P } from "../src/model";
import { getPartText, type DocxPackage, type DocxPart } from "../src/opc";

const enc = new TextEncoder();
const part = (path: string, content: string): [string, DocxPart] => [path, { path, bytes: enc.encode(content), dir: false }];

function pkg(): DocxPackage {
  return {
    order: [],
    parts: new Map([
      part(
        "word/document.xml",
        `<w:document><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p>` +
          `<w:sectPr><w:headerReference w:type="default" r:id="rId10"/><w:footerReference w:type="default" r:id="rId11"/></w:sectPr>` +
          `</w:body></w:document>`,
      ),
      part(
        "word/_rels/document.xml.rels",
        `<Relationships><Relationship Id="rId10" Target="header1.xml"/><Relationship Id="rId11" Target="footer1.xml"/></Relationships>`,
      ),
      part("word/header1.xml", `<w:hdr><w:p><w:r><w:t>ACME — Offer</w:t></w:r></w:p></w:hdr>`),
      part("word/footer1.xml", `<w:ftr><w:p><w:r><w:t>Confidential</w:t></w:r></w:p></w:ftr>`),
    ]),
  };
}

describe("headerFooter: resolve + parse", () => {
  const p = pkg();
  const model = parseDocument(getPartText(p, "word/document.xml")!);

  it("resolves the section's header/footer part XML", () => {
    expect(headerXml(p, model.section)).toContain("ACME — Offer");
    expect(footerXml(p, model.section)).toContain("Confidential");
  });

  it("parseContainer parses header/footer blocks", () => {
    const hdr = parseContainer(headerXml(p, model.section)!, "w:hdr");
    expect(paragraphText(hdr[0] as P)).toBe("ACME — Offer");
    const ftr = parseContainer(footerXml(p, model.section)!, "w:ftr");
    expect(paragraphText(ftr[0] as P)).toBe("Confidential");
  });

  it("returns undefined when no header reference", () => {
    expect(headerXml(p, undefined)).toBeUndefined();
  });
});

