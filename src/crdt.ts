// CRDT collaboration (ARCHITECTURE.md §3.15, §4b — SuperDoc uses Yjs + y-prosemirror; we bind Yjs to
// our string model instead). The document.xml lives in a Yjs Y.Text; a local edit is turned into a
// minimal text delta (common-prefix/suffix trim) and applied as Y.Text ops, so Yjs merges concurrent
// edits to different regions conflict-free — no clobber, unlike snapshot last-writer-wins. Updates are
// exchanged over the same pluggable transport (BroadcastChannel multi-tab, or a server for multi-client).
import * as Y from "yjs";
import type { CollabTransport } from "./collab";

/** Minimal single-span edit between two strings (common prefix + suffix trimmed). */
export function stringDelta(a: string, b: string): { index: number; remove: number; insert: string } {
  if (a === b) return { index: 0, remove: 0, insert: "" };
  const max = Math.min(a.length, b.length);
  let p = 0;
  while (p < max && a.charCodeAt(p) === b.charCodeAt(p)) p++;
  let s = 0;
  while (s < max - p && a.charCodeAt(a.length - 1 - s) === b.charCodeAt(b.length - 1 - s)) s++;
  return { index: p, remove: a.length - p - s, insert: b.slice(p, b.length - s) };
}

const b64 = {
  enc(u: Uint8Array): string {
    let bin = "";
    const C = 0x8000;
    for (let i = 0; i < u.length; i += C) bin += String.fromCharCode(...u.subarray(i, i + C));
    return typeof btoa !== "undefined" ? btoa(bin) : Buffer.from(u).toString("base64");
  },
  dec(str: string): Uint8Array {
    const bin = typeof atob !== "undefined" ? atob(str) : Buffer.from(str, "base64").toString("binary");
    const u = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return u;
  },
};

/** A peer's live presence (cursor + identity) for remote-cursor rendering. */
export interface PeerPresence {
  name?: string;
  color?: string;
  cursor?: { index: number; start: number; end: number };
}

export interface XmlCrdt {
  /** This peer's Yjs client id (stable per session) — handy for a presence colour. */
  clientId: number;
  /** Push a new local document.xml; diffed into a CRDT text op and broadcast. */
  applyLocal(documentXml: string): void;
  /** Subscribe to the merged document.xml (from any peer). Returns unsubscribe. */
  onChange(cb: (documentXml: string) => void): () => void;
  /** Broadcast this peer's presence (cursor/identity). */
  setPresence(state: PeerPresence): void;
  /** Subscribe to peer presence updates (map keyed by client id, excludes self). */
  onPresence(cb: (peers: Map<number, PeerPresence>) => void): () => void;
  current(): string;
  destroy(): void;
}

export function createXmlCrdt(initialXml: string, transport: CollabTransport, opts: { seedDelayMs?: number } = {}): XmlCrdt {
  const ydoc = new Y.Doc();
  const ytext = ydoc.getText("documentXml");
  let lastKnown = "";
  let seeded = false;
  const subs = new Set<(xml: string) => void>();
  const peers = new Map<number, PeerPresence>();
  const presenceSubs = new Set<(peers: Map<number, PeerPresence>) => void>();

  ydoc.on("update", (update: Uint8Array, origin: unknown) => {
    if (origin === "remote") return; // applied remote updates are not rebroadcast
    transport.send(JSON.stringify({ t: "u", d: b64.enc(update) }));
  });

  ytext.observe(() => {
    if (ytext.length > 0) seeded = true;
    const xml = ytext.toString();
    if (xml === lastKnown) return; // echo of our own local apply
    lastKnown = xml;
    for (const cb of [...subs]) cb(xml);
  });

  const off = transport.onReceive((msg) => {
    let m: { t?: string; d?: string; c?: number; s?: PeerPresence };
    try { m = JSON.parse(msg); } catch { return; }
    if (m.t === "p" && typeof m.c === "number") {
      if (m.c === ydoc.clientID) return;
      peers.set(m.c, m.s ?? {});
      for (const cb of [...presenceSubs]) cb(new Map(peers));
      return;
    }
    if (!m.d) return;
    if (m.t === "sv") transport.send(JSON.stringify({ t: "u", d: b64.enc(Y.encodeStateAsUpdate(ydoc, b64.dec(m.d))) }));
    else if (m.t === "u") Y.applyUpdate(ydoc, b64.dec(m.d), "remote");
  });

  // Ask peers for their state; if none seeds us first, seed from initialXml (jittered to avoid a
  // simultaneous double-seed across tabs; real multi-client deployments seed from the server).
  transport.send(JSON.stringify({ t: "sv", d: b64.enc(Y.encodeStateVector(ydoc)) }));
  const seedTimer = setTimeout(() => {
    if (!seeded && ytext.length === 0 && initialXml) {
      seeded = true;
      ydoc.transact(() => ytext.insert(0, initialXml));
    }
  }, (opts.seedDelayMs ?? 250) + Math.random() * 200);

  lastKnown = ytext.toString();

  return {
    applyLocal(documentXml) {
      if (documentXml === ytext.toString()) return;
      const d = stringDelta(ytext.toString(), documentXml);
      seeded = true;
      lastKnown = documentXml; // guard the observe echo
      ydoc.transact(() => {
        if (d.remove) ytext.delete(d.index, d.remove);
        if (d.insert) ytext.insert(d.index, d.insert);
      });
    },
    clientId: ydoc.clientID,
    onChange(cb) { subs.add(cb); return () => subs.delete(cb); },
    setPresence(state) { transport.send(JSON.stringify({ t: "p", c: ydoc.clientID, s: state })); },
    onPresence(cb) { presenceSubs.add(cb); return () => presenceSubs.delete(cb); },
    current() { return ytext.toString(); },
    destroy() { clearTimeout(seedTimer); off(); transport.close(); ydoc.destroy(); },
  };
}
