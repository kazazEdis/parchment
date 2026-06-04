import { describe, it, expect } from "vitest";
import { stringDelta, createXmlCrdt, type XmlCrdt, type PeerPresence } from "../src/crdt";
import { memoryTransport, type CollabTransport } from "../src/collab";

describe("crdt: stringDelta (minimal single-span diff)", () => {
  it("trims common prefix + suffix", () => {
    expect(stringDelta("abcXYZdef", "abcQQdef")).toEqual({ index: 3, remove: 3, insert: "QQ" });
  });
  it("pure insert / pure delete / no-op", () => {
    expect(stringDelta("abc", "abXYZc")).toEqual({ index: 2, remove: 0, insert: "XYZ" });
    expect(stringDelta("abcdef", "abef")).toEqual({ index: 2, remove: 2, insert: "" });
    expect(stringDelta("same", "same")).toEqual({ index: 0, remove: 0, insert: "" });
  });
});

/** A transport that queues messages until flushed — lets us simulate truly concurrent edits. */
function queueTransport() {
  const subs = new Set<(m: string) => void>();
  const q: string[] = [];
  const t: CollabTransport & { flush(): void } = {
    send(m) { q.push(m); },
    onReceive(cb) { subs.add(cb); return () => subs.delete(cb); },
    close() { subs.clear(); },
    flush() { for (const m of q.splice(0)) for (const cb of [...subs]) cb(m); },
  };
  return t;
}

describe("crdt: Yjs over document.xml", () => {
  const noSeed = { seedDelayMs: 1_000_000 };
  const drain = (t: { flush(): void }) => { for (let i = 0; i < 5; i++) t.flush(); };

  it("syncs an edit from one peer to another", () => {
    const t = queueTransport();
    const a = createXmlCrdt("", t, noSeed);
    const b = createXmlCrdt("", t, noSeed);
    drain(t);
    a.applyLocal("<w:body>hello</w:body>");
    drain(t);
    expect(b.current()).toBe("<w:body>hello</w:body>");
    a.destroy(); b.destroy();
  });

  it("merges CONCURRENT edits to different regions — no clobber (the LWW win)", () => {
    const t = queueTransport();
    const a = createXmlCrdt("", t, noSeed);
    const b = createXmlCrdt("", t, noSeed);
    drain(t);
    a.applyLocal("AAA BBB CCC");
    drain(t); // both peers now hold the same base
    expect(b.current()).toBe("AAA BBB CCC");

    // Concurrent: A edits the start, B edits the end — neither has seen the other yet.
    a.applyLocal("XXX BBB CCC"); // AAA → XXX
    b.applyLocal("AAA BBB YYY"); // CCC → YYY (b's base is still "AAA BBB CCC")
    drain(t);

    expect(a.current()).toBe(b.current()); // converged
    expect(a.current()).toBe("XXX BBB YYY"); // BOTH edits survived
    a.destroy(); b.destroy();
  });

  it("broadcasts presence to peers (excluding self)", () => {
    const t = memoryTransport();
    const a = createXmlCrdt("", t, noSeed);
    const b = createXmlCrdt("", t, noSeed);
    let bPeers = new Map<number, PeerPresence>();
    b.onPresence((p) => { bPeers = p; });
    a.setPresence({ name: "A", color: "#f00", cursor: { index: 0, start: 1, end: 3 } });
    expect(bPeers.get(a.clientId)).toEqual({ name: "A", color: "#f00", cursor: { index: 0, start: 1, end: 3 } });
    expect(bPeers.has(b.clientId)).toBe(false); // not self
    a.destroy(); b.destroy();
  });
});

