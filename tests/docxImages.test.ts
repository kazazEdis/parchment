import { describe, it, expect } from "vitest";
import { resolveImageDataUrl } from "../src/images";
import type { DocxPackage, DocxPart } from "../src/opc";

const enc = new TextEncoder();
const part = (path: string, content: string | Uint8Array): [string, DocxPart] => [
  path,
  { path, bytes: typeof content === "string" ? enc.encode(content) : content, dir: false },
];

function pkg(): DocxPackage {
  return {
    order: [],
    parts: new Map([
      part(
        "word/_rels/document.xml.rels",
        `<Relationships><Relationship Id="rId1" Target="media/image1.png"/><Relationship Id="rId2" Target="media/photo.jpeg"/></Relationships>`,
      ),
      part("word/media/image1.png", new Uint8Array([1, 2, 3])),
      part("word/media/photo.jpeg", new Uint8Array([255, 216, 255])),
    ]),
  };
}

describe("images: resolve blip relationship → data URL", () => {
  it("resolves a png by relationship id with correct MIME + base64", () => {
    expect(resolveImageDataUrl(pkg(), "rId1")).toBe("data:image/png;base64,AQID");
  });

  it("maps the extension to MIME (jpeg)", () => {
    expect(resolveImageDataUrl(pkg(), "rId2")).toBe("data:image/jpeg;base64,/9j/");
  });

  it("returns undefined for an unknown relationship", () => {
    expect(resolveImageDataUrl(pkg(), "rId9")).toBeUndefined();
  });
});

