import { describe, it, expect } from "vitest";
import { insertImage } from "../src/imageInsert";
import { resolveImageDataUrl } from "../src/images";
import { getPart, getPartText, type DocxPackage, type DocxPart } from "../src/opc";

const enc = new TextEncoder();
const part = (path: string, content: string | Uint8Array): [string, DocxPart] => [
  path,
  { path, bytes: typeof content === "string" ? enc.encode(content) : content, dir: false },
];

function pkg(): DocxPackage {
  return {
    order: [],
    parts: new Map([
      part("[Content_Types].xml", `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`),
      part("word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="styles.xml"/></Relationships>`),
    ]),
  };
}

describe("imageInsert", () => {
  it("registers media part + relationship + content-type and yields a drawing run", () => {
    const p = pkg();
    const { rId, runXml } = insertImage(p, { bytes: new Uint8Array([137, 80, 78, 71]), ext: "png", widthEmu: 914400, heightEmu: 457200 });

    expect(rId).toBe("rId2"); // next after rId1
    expect(getPart(p, "word/media/image1.png")).toBeTruthy();
    expect(getPartText(p, "[Content_Types].xml")).toContain('Extension="png"');
    expect(getPartText(p, "word/_rels/document.xml.rels")).toContain("media/image1.png");
    expect(runXml).toContain(`r:embed="rId2"`);
    expect(runXml).toContain(`cx="914400"`);

    // round-trips back to a data URL for rendering
    expect(resolveImageDataUrl(p, rId)).toMatch(/^data:image\/png;base64,/);
  });

  it("picks the next free media filename", () => {
    const p = pkg();
    insertImage(p, { bytes: new Uint8Array([1]), ext: "png", widthEmu: 100, heightEmu: 100 });
    insertImage(p, { bytes: new Uint8Array([2]), ext: "png", widthEmu: 100, heightEmu: 100 });
    expect(getPart(p, "word/media/image1.png")).toBeTruthy();
    expect(getPart(p, "word/media/image2.png")).toBeTruthy();
  });
});

