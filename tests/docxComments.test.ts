import { describe, it, expect } from "vitest";
import { parseComments, commentRanges, nextCommentId, addComment, wrapParagraphComment, wrapCommentRange } from "../src/comments";
import { parseDocument, paragraphText, type Paragraph as P } from "../src/model";

const COMMENTS = `<w:comments>
  <w:comment w:id="0" w:author="Ana Horvat" w:initials="AH" w:date="2026-06-04T00:00:00Z">
    <w:p><w:r><w:t xml:space="preserve">Please confirm </w:t></w:r><w:r><w:t>this total.</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="1" w:author="Bob"><w:p><w:r><w:t>Typo here</w:t></w:r></w:p></w:comment>
</w:comments>`;

const BODY = `<w:document><w:body><w:p>` +
  `<w:commentRangeStart w:id="0"/><w:r><w:t>Total: 1.234,56 EUR</w:t></w:r><w:commentRangeEnd w:id="0"/>` +
  `<w:r><w:commentReference w:id="0"/></w:r>` +
  `</w:p></w:body></w:document>`;

describe("comments: parse comments.xml", () => {
  it("reads id, author, initials, date, and flattened text", () => {
    const cs = parseComments(COMMENTS);
    expect(cs).toHaveLength(2);
    expect(cs[0]).toMatchObject({ id: "0", author: "Ana Horvat", initials: "AH", date: "2026-06-04T00:00:00Z" });
    expect(cs[0].text).toBe("Please confirm this total.");
    expect(cs[1]).toMatchObject({ id: "1", author: "Bob", text: "Typo here" });
  });

  it("handles a missing comments part", () => {
    expect(parseComments(undefined)).toEqual([]);
  });
});

describe("comments: locate commented ranges", () => {
  it("maps a comment id to the text it annotates", () => {
    expect(commentRanges(BODY).get("0")).toBe("Total: 1.234,56 EUR");
  });
});

describe("comments: authoring (write → parse back)", () => {
  it("nextCommentId returns the next free id", () => {
    expect(nextCommentId(COMMENTS)).toBe("2"); // ids 0,1 present
    expect(nextCommentId(undefined)).toBe("0");
  });

  it("addComment creates the part / appends, parseable round-trip", () => {
    const created = addComment(undefined, { id: "0", author: "Me", text: "First note" });
    expect(parseComments(created)).toEqual([{ id: "0", author: "Me", initials: undefined, date: undefined, text: "First note" }]);
    const appended = addComment(COMMENTS, { id: "2", author: "X", text: "Third" });
    const all = parseComments(appended);
    expect(all).toHaveLength(3);
    expect(all[2]).toMatchObject({ id: "2", text: "Third" });
  });

  it("wrapParagraphComment anchors a comment range in the body", () => {
    const xml = `<w:document><w:body><w:p><w:pPr><w:pStyle w:val="N"/></w:pPr><w:r><w:t>Total price line</w:t></w:r></w:p></w:body></w:document>`;
    const p = parseDocument(xml).body[0] as P;
    const wrapped = wrapParagraphComment(xml, p, "0");
    expect(commentRanges(wrapped).get("0")).toBe("Total price line"); // range now annotated
    expect(paragraphText(parseDocument(wrapped).body[0] as P)).toBe("Total price line"); // still parses cleanly
  });

  it("wrapCommentRange anchors a SUB-paragraph range (splits runs at the boundary)", () => {
    const xml = `<w:document><w:body><w:p><w:r><w:t>Hello World here</w:t></w:r></w:p></w:body></w:document>`;
    const p = parseDocument(xml).body[0] as P;
    const wrapped = wrapCommentRange(xml, p, 6, 11, "0"); // "World"
    expect(commentRanges(wrapped).get("0")).toBe("World");
    expect(paragraphText(parseDocument(wrapped).body[0] as P)).toBe("Hello World here");
  });
});

