import { describe, it, expect } from "vitest";
import { wrapHyperlink } from "../src/linkEdit";
import { parseDocument, paragraphText, type Paragraph as P, type Hyperlink } from "../src/model";

describe("linkEdit: wrapHyperlink", () => {
  it("wraps a sub-paragraph selection in a hyperlink", () => {
    const xml = `<w:document><w:body><w:p><w:r><w:t>Click here now</w:t></w:r></w:p></w:body></w:document>`;
    const p = parseDocument(xml).body[0] as P;
    const out = wrapHyperlink(xml, p, 6, 10, "rId5"); // "here"
    expect(out).toContain(`<w:hyperlink r:id="rId5">`);

    const rp = parseDocument(out).body[0] as P;
    const link = rp.children.find((n): n is Hyperlink => n.type === "hyperlink");
    expect(link?.rId).toBe("rId5");
    expect(paragraphText(rp)).toBe("Click here now"); // text intact
  });
});

