import { describe, it, expect } from "vitest";
import { memoryTransport } from "../src/collab";

describe("collab: snapshot transport", () => {
  it("broadcasts document.xml to subscribers and stops after unsubscribe", () => {
    const t = memoryTransport();
    const got: string[] = [];
    const off = t.onReceive((x) => got.push(x));
    t.send("<a/>");
    t.send("<b/>");
    expect(got).toEqual(["<a/>", "<b/>"]);
    off();
    t.send("<c/>");
    expect(got).toEqual(["<a/>", "<b/>"]); // no longer receiving
    t.close();
  });
});

