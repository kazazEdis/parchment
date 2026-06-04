import { describe, it, expect } from "vitest";
import { diffTokens, redlineRuns, redlineParagraph } from "../src/diff";
import { parseDocument, type Paragraph as P } from "../src/model";
import { emitParagraph } from "../src/serialize";

describe("diff: token diff", () => {
  it("coalesces an edit script", () => {
    expect(diffTokens(["a", "b", "c"], ["a", "x", "c"])).toEqual([
      { type: "eq", text: "a" },
      { type: "del", text: "b" },
      { type: "ins", text: "x" },
      { type: "eq", text: "c" },
    ]);
  });
});

describe("diff: redline runs", () => {
  it("marks word-level deletions and insertions with metadata", () => {
    const runs = redlineRuns("the quick fox", "the slow fox", { author: "A" });
    expect(runs.map((r) => r.text)).toEqual(["the ", "quick", "slow", " fox"]);
    expect(runs.map((r) => r.track?.type)).toEqual([undefined, "del", "ins", undefined]);
    expect(runs[1].track).toMatchObject({ type: "del", author: "A" });
  });

  it("redlineParagraph produces a paragraph that emits w:ins/w:del", () => {
    const p = parseDocument(`<w:document><w:body><w:p><w:r><w:t>old value here</w:t></w:r></w:p></w:body></w:document>`).body[0] as P;
    const xml = emitParagraph(redlineParagraph(p, "new value here", { author: "Rev" }));
    expect(xml).toContain("<w:del");
    expect(xml).toContain("<w:ins");
    expect(xml).toContain("value"); // unchanged words preserved as plain runs
  });
});

