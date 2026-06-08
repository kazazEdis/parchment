"use client";

// Interactive docx web editor (the WYSIWYG layer over the headless engine). Every edit is a pure
// engine op (edit/diff/serialize/doc) applied to the model, then re-rendered — so the editor inherits
// preserve-and-patch fidelity, determinism, and the engine's test coverage.
//
// Depth features beyond paragraph-level editing:
//  • inline range selection + per-selection formatting (formatRange — splits runs at the boundary)
//  • undo/redo (a reducer keeps the document + a history of document.xml snapshots — cheap + exact)
//  • paste from Word (text/html → runs, spliced at the caret via spliceRunRange)
// The robustness trick: on any structural op we rebuild the active paragraph from the LIVE DOM
// (capturing un-committed typing) and then apply the pure op — so model and view never desync.
//
// Paragraphs containing hyperlinks/drawings render read-only. Block-splitting (Enter→new paragraph)
// is still future. Not a full ProseMirror schema — but real inline editing, owned and light.
import React, { useCallback, useEffect, useMemo, useReducer, useRef, useState } from "react";
import { saveAs } from "file-saver";
import { type DocxPackage, setPartText } from "./opc";
import { parseComments } from "./comments";
import { broadcastChannelTransport } from "./collab";
import { createXmlCrdt, type XmlCrdt, type PeerPresence } from "./crdt";
import { parseDocument, type Block, type Paragraph, type Run, type Inline } from "./model";
import { type StyleSheet } from "./styles";
import { type Numbering } from "./numbering";
import { effectiveParagraphProps, effectiveRunProps, markerRunProps, assignListNumbers } from "./resolve";
import { paragraphCss, runCss, trackCss } from "./cssMap";
import { toggleBoolean, toggleUnderline, setAlignment, formatRange, spliceRunRange, rangeUniform, rangeUnderlined, splitParagraph, mergeParagraphs, setListLevel, paragraphLength } from "./edit";
import { redlineParagraph } from "./diff";
import { insertRowAfter, deleteRow, appendColumn, deleteColumn } from "./tableEdit";
import { insertImage } from "./imageInsert";
import { wrapHyperlink } from "./linkEdit";
import { ommlToMathML } from "./math";
import { addRelationship } from "./opcParts";
import type { Table } from "./model";
import { patchParagraph, patchSpan, patchAll, emitParagraph } from "./serialize";
import { twipsToPx } from "./units";
import { fromPackage, allParagraphs, replaceText, acceptAllChanges, rejectAllChanges, addCommentToParagraph, addCommentToRange, saveBlob, type Doc } from "./doc";

const isPureRuns = (p: Paragraph): boolean => p.children.length > 0 && p.children.every((n) => n.type === "run");
const renderRunText = (text: string): string => text.replace(/\t/g, "    ");

// Scoped toolbar styling (real :hover / :active / :disabled states, uniform icon buttons, grouped
// clusters). Kept as one injected stylesheet so the package stays dependency-free (no CSS file to
// import) while looking like a proper editor chrome rather than a flat row of beige buttons.
const PML_TOOLBAR_CSS = `
.pml-editor{font:13px/1.3 system-ui,-apple-system,Segoe UI,Roboto,sans-serif;color:#2b2b2b}
.pml-tb{position:sticky;top:0;z-index:5;display:flex;flex-wrap:wrap;align-items:center;gap:3px;padding:6px 10px;background:#fff;border-bottom:1px solid #e6e1d6;box-shadow:0 1px 0 rgba(0,0,0,.02)}
.pml-grp{display:flex;align-items:center;gap:2px}
.pml-sep{width:1px;align-self:stretch;margin:3px 5px;background:#e6e1d6}
.pml-btn{display:inline-flex;align-items:center;justify-content:center;gap:5px;min-width:32px;height:32px;padding:0 8px;border:1px solid transparent;background:transparent;border-radius:7px;cursor:pointer;color:#2b2b2b;font:inherit;line-height:1;transition:background .12s,border-color .12s,box-shadow .12s}
.pml-btn:hover:not(:disabled){background:#F4F0E7}
.pml-btn:active:not(:disabled){background:#e8e2d4}
.pml-btn:focus-visible{outline:none;box-shadow:0 0 0 2px rgba(28,87,66,.35)}
.pml-btn:disabled{opacity:.35;cursor:default}
.pml-ico{font-size:16px;line-height:1}
.pml-text{padding:0 10px;font-size:12.5px;font-weight:500;color:#3a3a3a}
.pml-b{font-weight:700;font-size:14px}
.pml-i{font-style:italic;font-family:Georgia,serif;font-size:14px}
.pml-u{text-decoration:underline;font-size:14px}
.pml-ai{font-weight:600;color:#1C5742}
.pml-ai:hover:not(:disabled){background:#1C57420f}
.pml-field{height:32px;border:1px solid #d8d2c4;border-radius:7px;padding:0 9px;font:inherit;background:#fff;color:#2b2b2b}
.pml-field::placeholder{color:#a9a193}
.pml-field:focus{outline:none;border-color:#1C5742;box-shadow:0 0 0 2px rgba(28,87,66,.12)}
.pml-chk{display:inline-flex;align-items:center;gap:5px;padding:0 7px;height:32px;color:#5b635f;white-space:nowrap;cursor:pointer}
.pml-chk input{cursor:pointer}
.pml-status{color:#8a8378;font-size:12px;margin-right:4px;white-space:nowrap}
.pml-primary{background:#1C5742;color:#fff;border-color:#1C5742;font-weight:600;padding:0 14px}
.pml-primary:hover:not(:disabled){background:#16442f}
.pml-primary:active:not(:disabled){background:#103625}
.pml-stage{background:#ECEAE3;padding:32px 24px;min-height:70vh;display:flex;gap:16px;justify-content:center}
`;

// ── live-DOM ↔ model helpers ──────────────────────────────────────────────────────────────────

/** Reconstruct a paragraph's runs from its contentEditable DOM (formatting-preserving). */
function runsFromDom(el: HTMLElement, original: Paragraph): Run[] {
  const originalRuns = original.children.filter((n): n is Run => n.type === "run");
  const out: Run[] = [];
  el.childNodes.forEach((node) => {
    if (node.nodeType === Node.ELEMENT_NODE && (node as HTMLElement).dataset.marker != null) return;
    const text = node.textContent ?? "";
    if (!text) return;
    let rPr: Run["rPr"] = {};
    let track: Run["track"];
    if (node.nodeType === Node.ELEMENT_NODE) {
      const ri = (node as HTMLElement).dataset.ri;
      const r = ri != null ? originalRuns[Number(ri)] : undefined;
      if (r) { rPr = r.rPr; track = r.track; }
      else if (out.length) { rPr = out[out.length - 1].rPr; track = out[out.length - 1].track; }
    } else if (out.length) { rPr = out[out.length - 1].rPr; track = out[out.length - 1].track; }
    else if (originalRuns[0]) rPr = originalRuns[0].rPr;
    const last = out[out.length - 1];
    if (last && last.rPr === rPr && last.track === track) last.text += text;
    else { const run: Run = { type: "run", rPr, text }; if (track) run.track = track; out.push(run); }
  });
  return out.length ? out : [{ type: "run", rPr: originalRuns[0]?.rPr ?? {}, text: "" }];
}

/** Character offset of a (node, offset) DOM position within a paragraph's run text (markers excluded). */
function caretOffset(paraEl: HTMLElement, container: Node, offset: number): number {
  let total = 0;
  const spans = Array.from(paraEl.querySelectorAll<HTMLElement>("span[data-ri]"));
  for (const span of spans) {
    const tn = span.firstChild;
    const len = span.textContent?.length ?? 0;
    if (container === tn) return total + offset;
    if (container === span) return total + (offset > 0 ? len : 0);
    if (span.contains(container)) return total + offset;
    total += len;
  }
  return total;
}

interface Sel { index: number; start: number; end: number }

/** The current selection mapped to {paragraphIndex, start, end}, or null if outside an editable paragraph. */
function currentSelection(): Sel | null {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  if (!sel || sel.rangeCount === 0) return null;
  const r = sel.getRangeAt(0);
  let node: Node | null = r.startContainer;
  while (node && !(node.nodeType === Node.ELEMENT_NODE && (node as HTMLElement).dataset?.pi != null)) node = node.parentNode;
  if (!node) return null;
  const paraEl = node as HTMLElement;
  if (!paraEl.contains(r.endContainer)) return null;
  const a = caretOffset(paraEl, r.startContainer, r.startOffset);
  const b = caretOffset(paraEl, r.endContainer, r.endOffset);
  return { index: Number(paraEl.dataset.pi), start: Math.min(a, b), end: Math.max(a, b) };
}

/** DOM (node, offset) for a character offset within a paragraph element's run spans. */
function locateOffset(paraEl: HTMLElement, offset: number): { node: Node; off: number } {
  let total = 0;
  for (const span of Array.from(paraEl.querySelectorAll<HTMLElement>("span[data-ri]"))) {
    const tn = span.firstChild ?? span;
    const len = span.textContent?.length ?? 0;
    if (offset <= total + len) return { node: tn, off: offset - total };
    total += len;
  }
  return { node: paraEl, off: paraEl.childNodes.length };
}

/** Place the selection at [start, end) within paragraph element `paraEl`. */
function restoreSelection(paraEl: HTMLElement, start: number, end: number): void {
  try {
    const a = locateOffset(paraEl, start);
    const b = locateOffset(paraEl, end);
    const range = document.createRange();
    range.setStart(a.node, a.off);
    range.setEnd(b.node, b.off);
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
  } catch {
    /* offsets out of range after a structural change — ignore */
  }
}

/** Parse pasted text/html (Word puts rich HTML on the clipboard) into runs; falls back to plain text. */
function htmlToRuns(html: string, baseRPr: Run["rPr"]): Run[] {
  if (typeof DOMParser === "undefined") return [{ type: "run", rPr: baseRPr, text: html.replace(/<[^>]+>/g, "") }];
  const doc = new DOMParser().parseFromString(html, "text/html");
  const runs: Run[] = [];
  const walk = (node: Node, rPr: Run["rPr"]): void => {
    node.childNodes.forEach((child) => {
      if (child.nodeType === Node.TEXT_NODE) {
        const text = (child.textContent ?? "").replace(/\s+/g, " ");
        if (text) runs.push({ type: "run", rPr, text });
        return;
      }
      if (child.nodeType !== Node.ELEMENT_NODE) return;
      const el = child as HTMLElement;
      const tag = el.tagName.toLowerCase();
      const r: Run["rPr"] = { ...rPr };
      if (tag === "b" || tag === "strong") r.bold = true;
      if (tag === "i" || tag === "em") r.italic = true;
      if (tag === "u") r.underline = "single";
      if (tag === "s" || tag === "strike" || tag === "del") r.strike = true;
      const st = el.style;
      if (st.fontWeight === "bold" || Number(st.fontWeight) >= 600) r.bold = true;
      if (st.fontStyle === "italic") r.italic = true;
      if ((st.textDecorationLine || st.textDecoration || "").includes("underline")) r.underline = "single";
      walk(el, r);
      if (tag === "br") runs.push({ type: "run", rPr, text: "\n" });
      else if (tag === "p" || tag === "div") runs.push({ type: "run", rPr, text: " " });
    });
  };
  walk(doc.body, baseRPr);
  return runs.length ? runs : [{ type: "run", rPr: baseRPr, text: "" }];
}

// ── reducer: document + undo/redo history (snapshots of document.xml) ─────────────────────────────

interface EState { doc: Doc; past: string[]; future: string[] }
type EAction =
  | { type: "commit"; produce: (d: Doc) => Doc }
  | { type: "undo" }
  | { type: "redo" }
  | { type: "remote"; xml: string };

function restoreXml(doc: Doc, xml: string): Doc {
  setPartText(doc.pkg, "word/document.xml", xml);
  return { ...doc, documentXml: xml, model: parseDocument(xml) };
}

function reducer(s: EState, a: EAction): EState {
  switch (a.type) {
    case "commit": {
      const nd = a.produce(s.doc);
      if (nd === s.doc) return s;
      return { doc: nd, past: [...s.past, s.doc.documentXml].slice(-100), future: [] };
    }
    case "undo": {
      if (!s.past.length) return s;
      const xml = s.past[s.past.length - 1];
      return { doc: restoreXml(s.doc, xml), past: s.past.slice(0, -1), future: [s.doc.documentXml, ...s.future] };
    }
    case "redo": {
      if (!s.future.length) return s;
      const xml = s.future[0];
      return { doc: restoreXml(s.doc, xml), past: [...s.past, s.doc.documentXml], future: s.future.slice(1) };
    }
    case "remote": {
      if (a.xml === s.doc.documentXml) return s;
      return { ...s, doc: restoreXml(s.doc, a.xml) };
    }
  }
}

/** Find the table (+ row index) that contains a paragraph, for table-structure edits. */
function findTableContext(body: Block[], paragraph: Paragraph): { table: Table; rowIndex: number } | null {
  for (const b of body) {
    if (b.type !== "table") continue;
    if (paragraph.source.start >= b.source.start && paragraph.source.end <= b.source.end) {
      let rowIndex = 0;
      b.rows.forEach((row, ri) => { if (paragraph.source.start >= row.source.start && paragraph.source.end <= row.source.end) rowIndex = ri; });
      return { table: b, rowIndex };
    }
  }
  return null;
}

/** Append a raw run's XML before a paragraph's closing tag. */
function appendRunToParagraph(documentXml: string, p: Paragraph, runXml: string): string {
  const outer = documentXml.slice(p.source.start, p.source.end);
  const tail = "</w:p>";
  if (!outer.endsWith(tail)) return documentXml;
  const newOuter = outer.slice(0, outer.length - tail.length) + runXml + tail;
  return documentXml.slice(0, p.source.start) + newOuter + documentXml.slice(p.source.end);
}

export function DocxEditor({ initialPackage, collabChannel, onAiRewrite, onChange }: { initialPackage: DocxPackage; collabChannel?: string; onAiRewrite?: (selectedText: string, instruction: string) => Promise<string>; onChange?: (doc: Doc) => void }): React.ReactElement {
  const [state, dispatch] = useReducer(reducer, undefined, () => ({ doc: fromPackage(initialPackage), past: [], future: [] }));
  const { doc } = state;
  // Notify the host of the current document on every change (initial + each edit) — lets a parent
  // drive an external Save / live preview from the edited bytes (doc.ts `save`/`saveBlob`).
  useEffect(() => { onChange?.(doc); }, [doc, onChange]);
  const [activeIndex, setActiveIndex] = useState<number | null>(null);
  const [track, setTrack] = useState(false);
  const [find, setFind] = useState("");
  const [replace, setReplace] = useState("");
  const [status, setStatus] = useState("");
  const activeElRef = useRef<HTMLElement | null>(null);
  const pendingSel = useRef<Sel | null>(null);

  const ctx = useMemo(() => ({ sheet: doc.styles as StyleSheet, numbering: doc.numbering as Numbering, markers: assignListNumbers(doc.model, doc.numbering) }), [doc]);
  const comments = useMemo(() => parseComments(doc.commentsXml), [doc.commentsXml]);

  // Collaboration: CRDT (Yjs over document.xml) — concurrent edits merge conflict-free, echo-guarded.
  const crdtRef = useRef<XmlCrdt | null>(null);
  const lastRemote = useRef<string | null>(null);
  const initialXmlRef = useRef(doc.documentXml);
  const myColorRef = useRef("#1C5742");
  const [remotePeers, setRemotePeers] = useState<Map<number, PeerPresence>>(new Map());
  const canvasWrapRef = useRef<HTMLDivElement>(null);
  const [cursorRects, setCursorRects] = useState<{ id: number; color: string; name: string; top: number; left: number; height: number }[]>([]);

  useEffect(() => {
    if (!collabChannel) return;
    const crdt = createXmlCrdt(initialXmlRef.current, broadcastChannelTransport(collabChannel));
    crdtRef.current = crdt;
    myColorRef.current = `hsl(${(crdt.clientId * 47) % 360} 70% 45%)`;
    const off = crdt.onChange((xml) => { lastRemote.current = xml; dispatch({ type: "remote", xml }); });
    const offP = crdt.onPresence((peers) => setRemotePeers(new Map(peers)));
    return () => { off(); offP(); crdt.destroy(); crdtRef.current = null; setRemotePeers(new Map()); };
  }, [collabChannel]);
  useEffect(() => {
    if (crdtRef.current && doc.documentXml !== lastRemote.current) crdtRef.current.applyLocal(doc.documentXml);
  }, [doc.documentXml]);
  // Broadcast our cursor on selection change.
  useEffect(() => {
    if (!collabChannel) return;
    const onSel = (): void => { crdtRef.current?.setPresence({ name: "You", color: myColorRef.current, cursor: currentSelection() ?? undefined }); };
    document.addEventListener("selectionchange", onSel);
    return () => document.removeEventListener("selectionchange", onSel);
  }, [collabChannel]);
  // Position remote carets relative to the canvas.
  useEffect(() => {
    const wrap = canvasWrapRef.current;
    if (!wrap || remotePeers.size === 0) { setCursorRects([]); return; }
    const base = wrap.getBoundingClientRect();
    const out: typeof cursorRects = [];
    for (const [id, p] of remotePeers) {
      if (!p.cursor) continue;
      const el = wrap.querySelector<HTMLElement>(`[data-pi="${p.cursor.index}"]`);
      if (!el) continue;
      try {
        const pos = locateOffset(el, p.cursor.start);
        const range = document.createRange();
        range.setStart(pos.node, pos.off);
        range.collapse(true);
        const r = range.getBoundingClientRect();
        out.push({ id, color: p.color ?? "#888", name: p.name ?? "User", top: r.top - base.top, left: r.left - base.left, height: r.height || 16 });
      } catch { /* ignore */ }
    }
    setCursorRects(out);
  }, [remotePeers, doc]);

  // Restore a pending selection after a structural commit re-rendered the paragraph.
  useEffect(() => {
    const s = pendingSel.current;
    if (!s) return;
    pendingSel.current = null;
    const el = document.querySelector<HTMLElement>(`[data-pi="${s.index}"]`);
    if (el) { el.focus(); restoreSelection(el, s.start, s.end); }
  });

  const commit = useCallback((produce: (d: Doc) => Doc) => dispatch({ type: "commit", produce }), []);

  /** Build the active paragraph from the live DOM (captures un-committed typing), apply `op`, write back. */
  const editActiveParagraph = useCallback((index: number, el: HTMLElement, op: (live: Paragraph) => Paragraph) => {
    commit((d) => {
      const model = allParagraphs(d.model)[index];
      if (!model) return d;
      const live: Paragraph = { ...model, children: runsFromDom(el, model) };
      const next = op(live);
      const newXml = patchParagraph(d.documentXml, next);
      setPartText(d.pkg, "word/document.xml", newXml);
      return { ...d, documentXml: newXml, model: parseDocument(newXml) };
    });
  }, [commit]);

  // Commit typed text on blur (track-changes mode → diff into redlines).
  const commitBlur = useCallback((index: number, el: HTMLElement) => {
    editActiveParagraph(index, el, (live) => {
      if (!track) return live;
      const text = live.children.filter((n): n is Run => n.type === "run").map((r) => r.text).join("");
      return redlineParagraph(allParagraphs(doc.model)[index]!, text, { author: "You", date: new Date().toISOString() });
    });
  }, [editActiveParagraph, track, doc.model]);

  // Toolbar formatting: range-aware when there is a non-empty selection, else whole paragraph.
  const format = useCallback((rangeOp: (p: Paragraph, s: number, e: number) => Paragraph, wholeOp: (p: Paragraph) => Paragraph) => {
    const el = activeElRef.current;
    if (!el || activeIndex == null) return;
    const sel = currentSelection();
    if (sel && sel.index === activeIndex && sel.end > sel.start) {
      pendingSel.current = sel;
      editActiveParagraph(activeIndex, el, (live) => rangeOp(live, sel.start, sel.end));
    } else {
      editActiveParagraph(activeIndex, el, wholeOp);
    }
  }, [activeIndex, editActiveParagraph]);

  const onBold = () => format((p, s, e) => formatRange(p, s, e, { bold: !rangeUniform(p, s, e, "bold") }), (p) => toggleBoolean(p, "bold"));
  const onItalic = () => format((p, s, e) => formatRange(p, s, e, { italic: !rangeUniform(p, s, e, "italic") }), (p) => toggleBoolean(p, "italic"));
  const onUnderline = () => format((p, s, e) => formatRange(p, s, e, { underline: rangeUnderlined(p, s, e) ? "none" : "single" }), (p) => toggleUnderline(p));

  const onPaste = useCallback((index: number, el: HTMLElement, e: React.ClipboardEvent) => {
    e.preventDefault();
    const html = e.clipboardData.getData("text/html");
    const text = e.clipboardData.getData("text/plain");
    const model = allParagraphs(doc.model)[index];
    const base = model?.children.find((n): n is Run => n.type === "run")?.rPr ?? {};
    const runs = html ? htmlToRuns(html, base) : [{ type: "run" as const, rPr: base, text }];
    const sel = currentSelection() ?? { index, start: 0, end: 0 };
    const insertedLen = runs.reduce((a, r) => a + r.text.length, 0);
    pendingSel.current = { index, start: sel.start + insertedLen, end: sel.start + insertedLen };
    editActiveParagraph(index, el, (live) => spliceRunRange(live, sel.start, sel.end, runs));
  }, [doc.model, editActiveParagraph]);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Table structure edits operate on the table containing the active paragraph (pure string ops).
  const tableOp = useCallback((fn: (xml: string, t: Table, ri: number) => string) => {
    if (activeIndex == null) return;
    const idx = activeIndex;
    commit((d) => {
      const p = allParagraphs(d.model)[idx];
      if (!p) return d;
      const tc = findTableContext(d.model.body, p);
      if (!tc) return d;
      const newXml = fn(d.documentXml, tc.table, tc.rowIndex);
      if (newXml === d.documentXml) return d;
      setPartText(d.pkg, "word/document.xml", newXml);
      return { ...d, documentXml: newXml, model: parseDocument(newXml) };
    });
  }, [activeIndex, commit]);

  const onInsertImage = useCallback(async (file: File) => {
    if (activeIndex == null) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    const ext = (file.name.split(".").pop() ?? "png").toLowerCase();
    const cx = 2743200; // 3 inches in EMU
    let cy = Math.round(cx * 0.6);
    try { const bmp = await createImageBitmap(file); cy = Math.round(cx * (bmp.height / bmp.width)); } catch { /* keep default */ }
    const { runXml } = insertImage(doc.pkg, { bytes, ext, widthEmu: cx, heightEmu: cy, name: file.name }); // media/rel/CT: once, outside reducer
    const idx = activeIndex;
    commit((d) => {
      const p = allParagraphs(d.model)[idx];
      if (!p) return d;
      const newXml = appendRunToParagraph(d.documentXml, p, runXml);
      setPartText(d.pkg, "word/document.xml", newXml);
      return { ...d, documentXml: newXml, model: parseDocument(newXml) };
    });
  }, [activeIndex, commit, doc.pkg]);

  const onAddLink = useCallback(() => {
    if (activeIndex == null) return;
    const sel = currentSelection();
    if (!sel || sel.index !== activeIndex || sel.end <= sel.start) { window.alert("Select text to link first."); return; }
    const url = window.prompt("Link URL:");
    if (!url) return;
    const rId = addRelationship(doc.pkg, "word/_rels/document.xml.rels", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", url, "External"); // once, outside reducer
    const s = sel;
    commit((d) => {
      const p = allParagraphs(d.model)[s.index];
      if (!p) return d;
      const newXml = wrapHyperlink(d.documentXml, p, s.start, s.end, rId);
      setPartText(d.pkg, "word/document.xml", newXml);
      return { ...d, documentXml: newXml, model: parseDocument(newXml) };
    });
  }, [activeIndex, commit, doc.pkg]);

  // AI assist: rewrite the selection via a product-supplied callback (wire it to Claude). Honors
  // track-changes (the rewrite lands as runs the diff/accept-reject flow can redline).
  const onAi = useCallback(async () => {
    if (activeIndex == null || !onAiRewrite) return;
    const sel = currentSelection();
    const el = activeElRef.current;
    if (!el || !sel || sel.index !== activeIndex || sel.end <= sel.start) { window.alert("Select text for the AI to rewrite."); return; }
    const model = allParagraphs(doc.model)[sel.index];
    if (!model) return;
    const live: Paragraph = { ...model, children: runsFromDom(el, model) };
    const fullText = live.children.filter((n): n is Run => n.type === "run").map((r) => r.text).join("");
    const selectedText = fullText.slice(sel.start, sel.end);
    const instruction = window.prompt("AI instruction (e.g. “make it more formal”):");
    if (!instruction) return;
    let result: string;
    try { result = await onAiRewrite(selectedText, instruction); } catch { window.alert("AI request failed."); return; }
    if (result == null) return;
    const s = sel;
    const rPr = live.children.find((n): n is Run => n.type === "run")?.rPr ?? {};
    commit((d) => {
      const p = allParagraphs(d.model)[s.index];
      if (!p) return d;
      const np = spliceRunRange(p, s.start, s.end, [{ type: "run", rPr, text: result }]);
      const newXml = patchParagraph(d.documentXml, np);
      setPartText(d.pkg, "word/document.xml", newXml);
      return { ...d, documentXml: newXml, model: parseDocument(newXml) };
    });
  }, [activeIndex, onAiRewrite, doc.model, commit]);

  // Enter → split the active paragraph at the caret (deleting any selection first).
  const splitActive = useCallback(() => {
    const el = activeElRef.current;
    if (!el || activeIndex == null) return;
    const sel = currentSelection();
    const start = sel && sel.index === activeIndex ? sel.start : 0;
    const end = sel && sel.index === activeIndex ? sel.end : start;
    const idx = activeIndex;
    pendingSel.current = { index: idx + 1, start: 0, end: 0 };
    commit((d) => {
      const model = allParagraphs(d.model)[idx];
      if (!model) return d;
      const live: Paragraph = { ...model, children: runsFromDom(el, model) };
      const cleared = end > start ? spliceRunRange(live, start, end, []) : live;
      const [p1, p2] = splitParagraph(cleared, start);
      const newXml = patchSpan(d.documentXml, model.source, emitParagraph(p1) + emitParagraph(p2));
      setPartText(d.pkg, "word/document.xml", newXml);
      return { ...d, documentXml: newXml, model: parseDocument(newXml) };
    });
    setActiveIndex(idx + 1);
  }, [activeIndex, commit]);

  // Backspace at start of a paragraph → merge it into the previous one.
  const mergeActive = useCallback(() => {
    const el = activeElRef.current;
    if (!el || activeIndex == null || activeIndex === 0) return;
    const idx = activeIndex;
    const prevModel = allParagraphs(doc.model)[idx - 1];
    if (!prevModel) return;
    pendingSel.current = { index: idx - 1, start: paragraphLength(prevModel), end: paragraphLength(prevModel) };
    commit((d) => {
      const cur = allParagraphs(d.model)[idx];
      const prev = allParagraphs(d.model)[idx - 1];
      if (!cur || !prev || !(prev.source.end <= cur.source.start && cur.source.start - prev.source.end < 24)) return d;
      const merged = mergeParagraphs(prev, { ...cur, children: runsFromDom(el, cur) });
      const newXml = patchAll(d.documentXml, [
        { span: prev.source, xml: emitParagraph(merged) },
        { span: cur.source, xml: "" },
      ]);
      setPartText(d.pkg, "word/document.xml", newXml);
      return { ...d, documentXml: newXml, model: parseDocument(newXml) };
    });
    setActiveIndex(idx - 1);
  }, [activeIndex, commit, doc.model]);

  // Tab / Shift+Tab → indent / outdent (list level or left indent).
  const indentActive = useCallback((delta: number) => {
    const el = activeElRef.current;
    if (!el || activeIndex == null) return;
    pendingSel.current = currentSelection();
    editActiveParagraph(activeIndex, el, (live) => setListLevel(live, delta));
  }, [activeIndex, editActiveParagraph]);

  const onKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (e.ctrlKey || e.metaKey) {
      const k = e.key.toLowerCase();
      if (k === "z" && !e.shiftKey) { e.preventDefault(); dispatch({ type: "undo" }); }
      else if (k === "y" || (k === "z" && e.shiftKey)) { e.preventDefault(); dispatch({ type: "redo" }); }
      else if (k === "b") { e.preventDefault(); onBold(); }
      else if (k === "i") { e.preventDefault(); onItalic(); }
      else if (k === "u") { e.preventDefault(); onUnderline(); }
      return;
    }
    if (activeIndex == null || !activeElRef.current) return;
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); splitActive(); }
    else if (e.key === "Tab") { e.preventDefault(); indentActive(e.shiftKey ? -1 : 1); }
    else if (e.key === "Backspace") {
      const sel = currentSelection();
      if (sel && sel.index === activeIndex && sel.start === 0 && sel.end === 0 && activeIndex > 0) { e.preventDefault(); mergeActive(); }
    }
  }, [activeIndex, onBold, onItalic, onUnderline, splitActive, mergeActive, indentActive]);

  const doReplace = () => {
    if (!find) return;
    commit((d) => { const r = replaceText(d, find, replace); setStatus(r.count ? `Replaced ${r.count}` : "No matches"); return r.doc; });
  };
  const download = async () => saveAs(await saveBlob(doc), "document.docx");

  // ── render ──
  const sec = doc.model.section;
  const m = sec?.margins ?? {};
  const page = { width: twipsToPx(sec?.pageSize?.width ?? 11906), padding: `${twipsToPx(m.top ?? 1440)}px ${twipsToPx(m.right ?? 1440)}px ${twipsToPx(m.bottom ?? 1440)}px ${twipsToPx(m.left ?? 1440)}px` };
  let pIndex = -1;

  const renderInlineStatic = (n: Inline, paraPPr: Paragraph["pPr"], key: number): React.ReactNode => {
    if (n.type === "run") {
      const css = runCss(effectiveRunProps(ctx.sheet, ctx.numbering, paraPPr, n.rPr));
      return <span key={key} style={n.track ? { ...css, ...trackCss(n.track.type) } : css}>{renderRunText(n.text)}</span>;
    }
    if (n.type === "hyperlink") return <a key={key} style={{ color: "#0563C1", textDecoration: "underline" }}>{n.children.map((c, i) => renderInlineStatic(c, paraPPr, i))}</a>;
    if (n.type === "footnoteRef") return <sup key={key} style={{ color: "#1C5742", fontSize: "0.7em" }}>{n.id}</sup>;
    if (n.type === "math") return <span key={key} dangerouslySetInnerHTML={{ __html: ommlToMathML(n.omml) }} />;
    return null;
  };

  const renderParagraph = (p: Paragraph): React.ReactElement => {
    pIndex += 1;
    const index = pIndex;
    const eff = effectiveParagraphProps(ctx.sheet, ctx.numbering, p.pPr);
    const marker = ctx.markers.get(p);
    const markerEl = marker !== undefined ? <span data-marker contentEditable={false} style={{ ...runCss(markerRunProps(ctx.sheet, ctx.numbering, p.pPr)), userSelect: "none", marginRight: "0.4em" }}>{marker}</span> : null;
    const runs = p.children.filter((n): n is Run => n.type === "run");
    const hasText = runs.some((r) => r.text.length > 0);

    if (!isPureRuns(p)) {
      return <p key={index} style={{ margin: 0, ...paragraphCss(eff) }}>{markerEl}{p.children.length ? p.children.map((n, i) => renderInlineStatic(n, p.pPr, i)) : <br />}</p>;
    }
    return (
      <p
        key={index}
        data-pi={index}
        contentEditable
        suppressContentEditableWarning
        onFocus={(e) => { setActiveIndex(index); activeElRef.current = e.currentTarget; }}
        onBlur={(e) => commitBlur(index, e.currentTarget)}
        onPaste={(e) => onPaste(index, e.currentTarget, e)}
        style={{ margin: 0, outline: activeIndex === index ? "1px solid #1C5742" : "none", borderRadius: 2, ...paragraphCss(eff) }}
      >
        {markerEl}
        {hasText
          ? runs.filter((r) => r.text.length > 0).map((r, i) => {
              const css = runCss(effectiveRunProps(ctx.sheet, ctx.numbering, p.pPr, r.rPr));
              return <span key={i} data-ri={i} style={r.track ? { ...css, ...trackCss(r.track.type) } : css}>{renderRunText(r.text)}</span>;
            })
          : <br />}
      </p>
    );
  };

  const renderBlock = (b: Block, key: number): React.ReactElement => {
    if (b.type === "paragraph") return <React.Fragment key={key}>{renderParagraph(b)}</React.Fragment>;
    const totalTwips = b.grid.reduce((a, c) => a + c, 0);
    return (
      <table key={key} style={{ borderCollapse: "collapse", tableLayout: "fixed", width: totalTwips ? twipsToPx(totalTwips) : "100%", margin: "0 0 8px" }}>
        <tbody>
          {b.rows.map((row, ri) => (
            <tr key={ri}>{row.cells.map((cell, ci) => <td key={ci} style={{ border: "1px solid #b3b3b3", padding: "2px 5px", verticalAlign: "top" }}>{cell.blocks.map((cb, bi) => renderBlock(cb, bi))}</td>)}</tr>
          ))}
        </tbody>
      </table>
    );
  };

  const md = (e: React.MouseEvent) => e.preventDefault(); // keep selection/focus on toolbar click
  const sep = <span className="pml-sep" />;
  // Three crisp alignment glyphs (SVG) — far cleaner than the old ⯇/≡/⯈ arrows.
  const alignIcon = (which: "left" | "center" | "right") => {
    const rows: Record<typeof which, [number, number][]> = {
      left:   [[2, 13], [2, 9], [2, 12]],
      center: [[3, 13], [5, 11], [4, 12]],
      right:  [[3, 14], [7, 14], [4, 14]],
    } as any;
    return (
      <svg width="15" height="15" viewBox="0 0 16 16" aria-hidden="true">
        <g stroke="currentColor" strokeWidth="1.5" strokeLinecap="round">
          {rows[which].map(([x1, x2], i) => <line key={i} x1={x1} y1={4 + i * 4} x2={x2} y2={4 + i * 4} />)}
        </g>
      </svg>
    );
  };

  return (
    <div className="pml-editor" onKeyDown={onKeyDown}>
      <style>{PML_TOOLBAR_CSS}</style>
      <div className="pml-tb">
        <div className="pml-grp">
          <button className="pml-btn" onMouseDown={md} onClick={() => dispatch({ type: "undo" })} disabled={!state.past.length} title="Undo (Ctrl+Z)"><span className="pml-ico">↶</span></button>
          <button className="pml-btn" onMouseDown={md} onClick={() => dispatch({ type: "redo" })} disabled={!state.future.length} title="Redo (Ctrl+Y)"><span className="pml-ico">↷</span></button>
        </div>
        {sep}
        <div className="pml-grp">
          <button className="pml-btn pml-b" onMouseDown={md} onClick={onBold} title="Bold (Ctrl+B)">B</button>
          <button className="pml-btn pml-i" onMouseDown={md} onClick={onItalic} title="Italic (Ctrl+I)">I</button>
          <button className="pml-btn pml-u" onMouseDown={md} onClick={onUnderline} title="Underline (Ctrl+U)">U</button>
        </div>
        {sep}
        <div className="pml-grp">
          <button className="pml-btn" onMouseDown={md} onClick={() => format((p) => p, (p) => setAlignment(p, "left"))} title="Align left">{alignIcon("left")}</button>
          <button className="pml-btn" onMouseDown={md} onClick={() => format((p) => p, (p) => setAlignment(p, "center"))} title="Align center">{alignIcon("center")}</button>
          <button className="pml-btn" onMouseDown={md} onClick={() => format((p) => p, (p) => setAlignment(p, "right"))} title="Align right">{alignIcon("right")}</button>
        </div>
        {sep}
        <div className="pml-grp">
          <input className="pml-field" placeholder="Find" value={find} onChange={(e) => setFind(e.target.value)} style={{ width: 90 }} />
          <input className="pml-field" placeholder="Replace" value={replace} onChange={(e) => setReplace(e.target.value)} style={{ width: 90 }} />
          <button className="pml-btn pml-text" onClick={doReplace}>Replace</button>
        </div>
        {sep}
        <div className="pml-grp">
          <label className="pml-chk" title="Record edits as tracked changes"><input type="checkbox" checked={track} onChange={(e) => setTrack(e.target.checked)} /> Track</label>
          <button className="pml-btn pml-text" onClick={() => commit((d) => acceptAllChanges(d))} title="Accept all tracked changes">Accept</button>
          <button className="pml-btn pml-text" onClick={() => commit((d) => rejectAllChanges(d))} title="Reject all tracked changes">Reject</button>
          <button
            className="pml-btn"
            onMouseDown={md}
            title="Comment on the active paragraph"
            onClick={() => {
              if (activeIndex == null) return;
              const sel = currentSelection(); // capture before the modal prompt collapses it
              const text = window.prompt("Comment:");
              if (!text) return;
              if (sel && sel.index === activeIndex && sel.end > sel.start) {
                const s = sel;
                commit((d) => addCommentToRange(d, s.index, s.start, s.end, { author: "You", text }));
              } else {
                commit((d) => addCommentToParagraph(d, activeIndex, { author: "You", text }));
              }
            }}
          ><span className="pml-ico">🗨</span></button>
        </div>
        {sep}
        <div className="pml-grp">
          {onAiRewrite && <button className="pml-btn pml-ai" onMouseDown={md} onClick={() => void onAi()} title="AI: rewrite the selection">✦ AI</button>}
          <button className="pml-btn pml-text" onMouseDown={md} onClick={() => fileInputRef.current?.click()} title="Insert image into the active paragraph">Image</button>
          <button className="pml-btn pml-text" onMouseDown={md} onClick={onAddLink} title="Hyperlink the selection">Link</button>
        </div>
        {sep}
        <div className="pml-grp" title="Table (inside a table cell)">
          <button className="pml-btn pml-text" onMouseDown={md} onClick={() => tableOp(insertRowAfter)} title="Add a row">+Row</button>
          <button className="pml-btn pml-text" onMouseDown={md} onClick={() => tableOp(deleteRow)} title="Delete the row">−Row</button>
          <button className="pml-btn pml-text" onMouseDown={md} onClick={() => tableOp((x, t) => appendColumn(x, t))} title="Add a column">+Col</button>
          <button className="pml-btn pml-text" onMouseDown={md} onClick={() => tableOp((x, t) => deleteColumn(x, t, 0))} title="Delete the first column">−Col</button>
        </div>
        <input ref={fileInputRef} type="file" accept="image/*" style={{ display: "none" }} onChange={(e) => { const f = e.target.files?.[0]; if (f) void onInsertImage(f); e.target.value = ""; }} />
        <span style={{ flex: 1 }} />
        {status && <span className="pml-status">{status}</span>}
        <button className="pml-btn pml-primary" onClick={download} title="Download the edited .docx">Download .docx</button>
      </div>

      <div className="pml-stage">
        <div ref={canvasWrapRef} style={{ position: "relative" }}>
          <div id="docx-editor-canvas" style={{ width: page.width, background: "#fff", boxShadow: "0 2px 12px rgba(0,0,0,0.12)", borderRadius: 2, padding: page.padding, boxSizing: "border-box", fontFamily: "Calibri, Carlito, sans-serif", fontSize: "11pt", lineHeight: 1.15, color: "#000" }}>
            {doc.model.body.map((b, i) => renderBlock(b, i))}
          </div>
          {cursorRects.map((c) => (
            <div key={c.id} style={{ position: "absolute", top: c.top, left: c.left, pointerEvents: "none", zIndex: 4 }}>
              <div style={{ width: 2, height: c.height, background: c.color }} />
              <div style={{ position: "absolute", top: -13, left: 0, background: c.color, color: "#fff", fontSize: 9, lineHeight: "12px", padding: "0 3px", borderRadius: 2, whiteSpace: "nowrap" }}>{c.name}</div>
            </div>
          ))}
        </div>
        {comments.length > 0 && (
          <aside id="docx-comments" style={{ width: 240, fontSize: 12 }}>
            <div style={{ fontWeight: 600, color: "#5b635f", margin: "4px 0 8px" }}>Comments ({comments.length})</div>
            {comments.map((c) => (
              <div key={c.id} style={{ background: "#fff", border: "1px solid #e0dacd", borderRadius: 6, padding: "8px 10px", marginBottom: 8 }}>
                <div style={{ fontWeight: 600 }}>{c.author ?? "—"}</div>
                <div style={{ color: "#333", marginTop: 2 }}>{c.text}</div>
              </div>
            ))}
          </aside>
        )}
      </div>
    </div>
  );
}
