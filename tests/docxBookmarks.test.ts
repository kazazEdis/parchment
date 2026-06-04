import { describe, it, expect } from "vitest";
import { isEncryptedOfficeFile } from "../src/encrypted";
import { parseBookmarks, insertBookmark } from "../src/bookmarks";
import { parseDocument, paragraphText, type Paragraph as P } from "../src/model";

describe("encrypted: detect", () => {
  it("recognises the OLE/CFB magic; ZIP and short inputs are not encrypted", () => {
    expect(isEncryptedOfficeFile(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0]))).toBe(true);
    expect(isEncryptedOfficeFile(new Uint8Array([0x50, 0x4b, 0x03, 0x04]))).toBe(false); // "PK" zip
    expect(isEncryptedOfficeFile(new Uint8Array([1, 2]))).toBe(false);
  });
});

describe("bookmarks", () => {
  it("parses named bookmarks, filtering _GoBack", () => {
    const xml =
      `<w:document><w:body>` +
      `<w:p><w:bookmarkStart w:id="0" w:name="ref1"/><w:r><w:t>X</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>` +
      `<w:p><w:bookmarkStart w:id="1" w:name="_GoBack"/><w:bookmarkEnd w:id="1"/></w:p>` +
      `</w:body></w:document>`;
    expect(parseBookmarks(xml)).toEqual([{ id: "0", name: "ref1" }]);
  });

  it("inserts a bookmark over a range and round-trips", () => {
    const src = `<w:document><w:body><w:p><w:r><w:t>Hello World</w:t></w:r></w:p></w:body></w:document>`;
    const p = parseDocument(src).body[0] as P;
    const out = insertBookmark(src, p, 0, 5, "0", "greeting"); // "Hello"
    expect(parseBookmarks(out)).toEqual([{ id: "0", name: "greeting" }]);
    expect(paragraphText(parseDocument(out).body[0] as P)).toBe("Hello World");
  });
});

