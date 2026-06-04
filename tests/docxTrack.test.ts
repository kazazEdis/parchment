import { describe, it, expect } from "vitest";
import { parseDocument, paragraphText, type Paragraph as P, type Run } from "../src/model";
import { emitParagraph } from "../src/serialize";
import { acceptChanges, rejectChanges, markTracked } from "../src/edit";

const DOC =
  `<w:document><w:body><w:p>` +
  `<w:r><w:t xml:space="preserve">keep </w:t></w:r>` +
  `<w:ins w:id="1" w:author="Ana" w:date="2026-06-04T00:00:00Z"><w:r><w:t xml:space="preserve">added </w:t></w:r></w:ins>` +
  `<w:del w:id="2" w:author="Bob"><w:r><w:delText xml:space="preserve">gone </w:delText></w:r></w:del>` +
  `<w:r><w:t>end</w:t></w:r>` +
  `</w:p></w:body></w:document>`;

const runsOf = (p: P): Run[] => p.children.filter((n): n is Run => n.type === "run");

describe("track changes: model captures ins/del", () => {
  const p = parseDocument(DOC).body[0] as P;
  it("reads runs with revision metadata + delText", () => {
    const rs = runsOf(p);
    expect(rs.map((r) => r.text)).toEqual(["keep ", "added ", "gone ", "end"]);
    expect(rs[0].track).toBeUndefined();
    expect(rs[1].track).toEqual({ type: "ins", id: "1", author: "Ana", date: "2026-06-04T00:00:00Z" });
    expect(rs[2].track).toMatchObject({ type: "del", author: "Bob" });
  });
});

describe("track changes: serialize round-trip", () => {
  it("emits w:ins/w:del + w:delText and reparses identically", () => {
    const p = parseDocument(DOC).body[0] as P;
    const xml = emitParagraph(p);
    expect(xml).toContain("<w:ins ");
    expect(xml).toContain("<w:del ");
    expect(xml).toContain("<w:delText");
    const rp = parseDocument(`<w:document><w:body>${xml}</w:body></w:document>`).body[0] as P;
    expect(runsOf(rp).map((r) => ({ t: r.text, k: r.track?.type }))).toEqual(
      runsOf(p).map((r) => ({ t: r.text, k: r.track?.type })),
    );
  });
});

describe("track changes: accept / reject", () => {
  const p = parseDocument(DOC).body[0] as P;

  it("accept keeps insertions (as plain runs) and drops deletions", () => {
    const a = acceptChanges(p);
    expect(paragraphText(a)).toBe("keep added end");
    expect(runsOf(a).every((r) => r.track === undefined)).toBe(true);
  });

  it("reject drops insertions and restores deletions", () => {
    const r = rejectChanges(p);
    expect(paragraphText(r)).toBe("keep gone end");
    expect(runsOf(r).every((x) => x.track === undefined)).toBe(true);
  });

  it("markTracked wraps runs so they emit as a tracked insertion", () => {
    const base = parseDocument(`<w:document><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>`).body[0] as P;
    const xml = emitParagraph(markTracked(base, { type: "ins", author: "Me" }));
    expect(xml).toContain(`<w:ins w:id="0" w:author="Me">`);
  });
});

