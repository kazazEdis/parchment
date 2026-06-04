import { describe, it, expect } from "vitest";
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from "docx";
import {
  readDocx,
  writeDocx,
  getPartText,
  setPartText,
  hasPart,
} from "../src/opc";

// Build a real .docx fixture with the (MIT) `docx` lib so the test is self-contained.
async function fixtureDocx(): Promise<Uint8Array> {
  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Offer")] }),
          new Paragraph({ children: [new TextRun("Number: {offer_number}")] }),
          new Paragraph({ children: [new TextRun("Customer: {customer_name}")] }),
        ],
      },
    ],
  });
  const buf = await Packer.toBuffer(doc);
  return new Uint8Array(buf);
}

describe("opc: read/write .docx package", () => {
  it("reads the expected OOXML parts", async () => {
    const pkg = await readDocx(await fixtureDocx());
    expect(hasPart(pkg, "word/document.xml")).toBe(true);
    expect(hasPart(pkg, "[Content_Types].xml")).toBe(true);
    expect(hasPart(pkg, "_rels/.rels")).toBe(true);
    const body = getPartText(pkg, "word/document.xml") ?? "";
    expect(body).toContain("{offer_number}");
    expect(body).toContain("{customer_name}");
  });

  it("identity round-trip preserves every part byte-for-byte", async () => {
    const original = await readDocx(await fixtureDocx());
    const rezipped = await writeDocx(original);
    const reopened = await readDocx(rezipped);

    // Same set of parts, same order.
    expect(reopened.order).toEqual(original.order);

    // Each part's decompressed bytes are identical (compression may differ; content must not).
    for (const path of original.order) {
      const a = original.parts.get(path)!;
      const b = reopened.parts.get(path)!;
      expect(b, `missing part ${path}`).toBeTruthy();
      expect(b.dir).toBe(a.dir);
      expect(Array.from(b.bytes), `bytes differ for ${path}`).toEqual(Array.from(a.bytes));
    }
  });

  it("patching one part leaves the others untouched (preserve-and-patch foundation)", async () => {
    const pkg = await readDocx(await fixtureDocx());
    const before = new Map(
      pkg.order
        .filter((p) => p !== "word/document.xml")
        .map((p) => [p, Array.from(pkg.parts.get(p)!.bytes)] as const),
    );

    // Simulate an edit to just the body part.
    const body = getPartText(pkg, "word/document.xml")!;
    setPartText(pkg, "word/document.xml", body.replace("Offer", "Quotation"));
    expect(pkg.parts.get("word/document.xml")!.dirty).toBe(true);

    const reopened = await readDocx(await writeDocx(pkg));

    // Edited part changed...
    expect(getPartText(reopened, "word/document.xml")).toContain("Quotation");
    // ...every other part is byte-identical.
    for (const [path, bytes] of before) {
      expect(Array.from(reopened.parts.get(path)!.bytes), `part drifted: ${path}`).toEqual(bytes);
    }
  });
});

