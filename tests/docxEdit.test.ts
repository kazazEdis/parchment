import { describe, it, expect } from "vitest";
import { parseDocument, paragraphText, type Paragraph as P, type Run } from "../src/model";
import {
  formatRuns,
  withParagraphProps,
  setAlignment,
  toggleBoolean,
  toggleUnderline,
  applyEdit,
} from "../src/edit";

const DOC =
  `<w:document><w:body>` +
  `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">bold </w:t></w:r><w:r><w:t>plain</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>second</w:t></w:r></w:p>` +
  `</w:body></w:document>`;

const runsOf = (p: P): Run[] => p.children.filter((n): n is Run => n.type === "run");

describe("edit: pure paragraph transforms", () => {
  const p0 = parseDocument(DOC).body[0] as P;

  it("formatRuns applies a run-prop patch to every run", () => {
    const out = formatRuns(p0, { color: "00FF00" });
    expect(runsOf(out).map((r) => r.rPr.color)).toEqual(["00FF00", "00FF00"]);
    expect(p0.children).not.toBe(out.children); // immutable: original untouched
  });

  it("toggleBoolean turns the format on everywhere when not already uniform, then off", () => {
    const on = toggleBoolean(p0, "bold"); // run1 wasn't bold → make all bold
    expect(runsOf(on).map((r) => r.rPr.bold)).toEqual([true, true]);
    const off = toggleBoolean(on, "bold"); // all bold → unbold
    expect(runsOf(off).map((r) => r.rPr.bold)).toEqual([false, false]);
  });

  it("toggleUnderline flips none ⇄ single", () => {
    expect(runsOf(toggleUnderline(p0)).map((r) => r.rPr.underline)).toEqual(["single", "single"]);
  });

  it("setAlignment + withParagraphProps update pPr", () => {
    expect(setAlignment(p0, "center").pPr.alignment).toBe("center");
    expect(withParagraphProps(p0, { keepNext: true }).pPr.keepNext).toBe(true);
  });
});

describe("edit: applyEdit writes back via preserve-and-patch", () => {
  it("edits one paragraph and leaves the rest of document.xml intact", () => {
    const model = parseDocument(DOC);
    const p0 = model.body[0] as P;
    const out = applyEdit(DOC, p0, (p) => formatRuns(p, { italic: true }));

    const re = parseDocument(out);
    expect(runsOf(re.body[0] as P).every((r) => r.rPr.italic === true)).toBe(true);
    expect(runsOf(re.body[0] as P).map((r) => r.rPr.bold)).toEqual([true, undefined]); // bold preserved on run0
    expect(paragraphText(re.body[1] as P)).toBe("second"); // second paragraph untouched
    expect(out.slice(0, p0.source.start)).toBe(DOC.slice(0, p0.source.start)); // bytes before the edit intact
  });
});

