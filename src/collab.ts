// Collaboration layer (ARCHITECTURE.md §3.15). Honest scope: document.xml **snapshot sync** over a
// pluggable transport — every local change broadcasts the new document.xml; peers replace theirs.
// Multi-tab (BroadcastChannel) and multi-client (any send/onReceive transport, e.g. a Convex doc or
// WebSocket) work today. This is last-writer-wins, NOT a CRDT: simultaneous edits to different parts
// can clobber. True conflict-free concurrent editing is a Yjs binding on top of the model — the
// documented upgrade path. The transport abstraction is deliberately tiny so that swap is local.

export interface CollabTransport {
  /** Broadcast the latest document.xml to peers. */
  send(documentXml: string): void;
  /** Subscribe to peer updates; returns an unsubscribe function. */
  onReceive(cb: (documentXml: string) => void): () => void;
  /** Release resources. */
  close(): void;
}

/** Same-origin multi-tab transport via BroadcastChannel. No-ops where BroadcastChannel is absent. */
export function broadcastChannelTransport(channelName: string): CollabTransport {
  const ch = typeof BroadcastChannel !== "undefined" ? new BroadcastChannel(channelName) : null;
  return {
    send(documentXml) { ch?.postMessage({ documentXml }); },
    onReceive(cb) {
      if (!ch) return () => {};
      const handler = (e: MessageEvent) => { if (e.data?.documentXml != null) cb(e.data.documentXml as string); };
      ch.addEventListener("message", handler);
      return () => ch.removeEventListener("message", handler);
    },
    close() { ch?.close(); },
  };
}

/** In-memory transport (testing / single-process multi-instance). */
export function memoryTransport(): CollabTransport {
  const subs = new Set<(xml: string) => void>();
  return {
    send(documentXml) { for (const cb of [...subs]) cb(documentXml); },
    onReceive(cb) { subs.add(cb); return () => subs.delete(cb); },
    close() { subs.clear(); },
  };
}
