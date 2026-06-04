// Editing operations (ARCHITECTURE.md §3, step 7). Pure model→model transforms over a Paragraph;
// compose with serialize.patchParagraph to write the edit back into document.xml at the paragraph's
// source span (preserve-and-patch). Keeping edits pure + node-level makes them trivially testable and
// keeps the interactive UI a thin dispatcher over these functions.
import type { Paragraph, Run, Inline, TrackChange } from "./model";
import type { RunProps, ParagraphProps } from "./props";
import { mergeParagraphProps } from "./props";
import { patchParagraph } from "./serialize";

/** Map every run in the paragraph (descending into hyperlinks), leaving non-runs untouched. */
export function mapRuns(p: Paragraph, fn: (r: Run) => Run): Paragraph {
  const map = (nodes: Inline[]): Inline[] =>
    nodes.map((n) =>
      n.type === "run" ? fn(n) : n.type === "hyperlink" ? { ...n, children: map(n.children) } : n,
    );
  return { ...p, children: map(p.children) };
}

/** Apply a partial run-property patch to every run in the paragraph. */
export function formatRuns(p: Paragraph, patch: Partial<RunProps>): Paragraph {
  return mapRuns(p, (r) => ({ ...r, rPr: { ...r.rPr, ...patch } }));
}

/** Merge a paragraph-property patch (deep for indent/spacing). */
export function withParagraphProps(p: Paragraph, patch: ParagraphProps): Paragraph {
  return { ...p, pPr: mergeParagraphProps(p.pPr, patch) };
}

/** Set the paragraph alignment. */
export function setAlignment(p: Paragraph, alignment: ParagraphProps["alignment"]): Paragraph {
  return { ...p, pPr: { ...p.pPr, alignment } };
}

const allRuns = (p: Paragraph): Run[] => {
  const acc: Run[] = [];
  const walk = (nodes: Inline[]): void => {
    for (const n of nodes) {
      if (n.type === "run") acc.push(n);
      else if (n.type === "hyperlink") walk(n.children);
    }
  };
  walk(p.children);
  return acc;
};

/**
 * Toggle a boolean run format across the paragraph, Word-style: if every run already has it on, turn
 * it off everywhere; otherwise turn it on everywhere.
 */
export function toggleBoolean(p: Paragraph, key: "bold" | "italic" | "strike" | "caps" | "smallCaps"): Paragraph {
  const runs = allRuns(p);
  const allOn = runs.length > 0 && runs.every((r) => r.rPr[key] === true);
  return formatRuns(p, { [key]: !allOn } as Partial<RunProps>);
}

/** Toggle underline across the paragraph (single ⇄ none), Word-style. */
export function toggleUnderline(p: Paragraph): Paragraph {
  const runs = allRuns(p);
  const allOn = runs.length > 0 && runs.every((r) => r.rPr.underline !== undefined && r.rPr.underline !== "none");
  return formatRuns(p, { underline: allOn ? "none" : "single" });
}

/**
 * Replace every `search` with `replace` inside one paragraph, **correctly across run boundaries** —
 * Word routinely splits a string like "{customer_name}" across several runs (rsid/spellcheck), which
 * defeats naive per-run replacement (and is why template fills need a dedicated library). Replacement
 * text inherits the formatting of the run where the match begins. Paragraphs that contain non-run
 * inlines (hyperlinks/drawings) fall back to independent per-run replacement. Returns the new
 * paragraph and the replacement count.
 */
export function replaceInParagraph(p: Paragraph, search: string, replace: string): { paragraph: Paragraph; count: number } {
  if (!search) return { paragraph: p, count: 0 };

  if (!p.children.every((n) => n.type === "run")) {
    let count = 0;
    const paragraph = mapRuns(p, (r) => {
      const parts = r.text.split(search);
      if (parts.length > 1) { count += parts.length - 1; return { ...r, text: parts.join(replace) }; }
      return r;
    });
    return { paragraph, count };
  }

  type Cell = { ch: string; rPr: Run["rPr"]; track?: TrackChange };
  const chars: Cell[] = [];
  for (const r of p.children as Run[]) for (const ch of r.text) chars.push({ ch, rPr: r.rPr, track: r.track });
  const s = chars.map((c) => c.ch).join("");

  let count = 0;
  const result: Cell[] = [];
  let i = 0;
  while (i < s.length) {
    if (s.startsWith(search, i)) {
      const fmt = chars[i]; // formatting at the match start
      for (const ch of replace) result.push({ ch, rPr: fmt.rPr, track: fmt.track });
      i += search.length;
      count++;
    } else {
      result.push(chars[i]);
      i++;
    }
  }
  if (!count) return { paragraph: p, count: 0 };

  // Regroup consecutive chars sharing the same formatting reference back into runs.
  const out: Run[] = [];
  for (const c of result) {
    const last = out[out.length - 1];
    if (last && last.rPr === c.rPr && last.track === c.track) last.text += c.ch;
    else { const run: Run = { type: "run", rPr: c.rPr, text: c.ch }; if (c.track) run.track = c.track; out.push(run); }
  }
  return { paragraph: { ...p, children: out }, count };
}

// ── range editing (inline selection, the WYSIWYG primitives) ─────────────────────────────────────
// Offsets are character positions in the paragraph's flattened run text. These power per-selection
// formatting, paste, and delete; they automatically split/merge runs at the range boundaries.

interface CharCell {
  ch: string;
  rPr: RunProps;
  track?: TrackChange;
}

const clamp = (n: number, lo: number, hi: number): number => Math.max(lo, Math.min(n, hi));

/** Flatten a run-only paragraph to per-character cells; null if it has non-run inlines. */
function flattenChars(p: Paragraph): CharCell[] | null {
  if (!p.children.every((n) => n.type === "run")) return null;
  const out: CharCell[] = [];
  for (const r of p.children as Run[]) for (const ch of r.text) out.push({ ch, rPr: r.rPr, track: r.track });
  return out;
}

/** Coalesce adjacent cells sharing the same rPr/track reference back into runs. */
function runsFromCells(cells: CharCell[]): Run[] {
  const out: Run[] = [];
  for (const c of cells) {
    const last = out[out.length - 1];
    if (last && last.rPr === c.rPr && last.track === c.track) last.text += c.ch;
    else { const run: Run = { type: "run", rPr: c.rPr, text: c.ch }; if (c.track) run.track = c.track; out.push(run); }
  }
  return out;
}

/** The plain text length of a run-only paragraph (for clamping selection offsets). */
export function paragraphLength(p: Paragraph): number {
  return (p.children.filter((n): n is Run => n.type === "run")).reduce((a, r) => a + r.text.length, 0);
}

/**
 * Apply a run-property patch to the character range [start, end), splitting runs at the boundaries.
 * Chars from the same source run share the patched rPr (so they stay one run); chars from different
 * source runs stay distinct. No-op on paragraphs with non-run inlines, or empty ranges.
 */
export function formatRange(p: Paragraph, start: number, end: number, patch: Partial<RunProps>): Paragraph {
  const cells = flattenChars(p);
  if (!cells) return p;
  const s = clamp(start, 0, cells.length);
  const e = clamp(end, s, cells.length);
  if (s === e) return p;
  const cache = new Map<RunProps, RunProps>();
  const patched = (rPr: RunProps): RunProps => {
    let v = cache.get(rPr);
    if (!v) { v = { ...rPr, ...patch }; cache.set(rPr, v); }
    return v;
  };
  const next = cells.map((c, i) => (i >= s && i < e ? { ...c, rPr: patched(c.rPr) } : c));
  return { ...p, children: runsFromCells(next) };
}

/** Replace the character range [start, end) with `insert` runs (insert/delete/replace at a caret). */
export function spliceRunRange(p: Paragraph, start: number, end: number, insert: Run[]): Paragraph {
  const cells = flattenChars(p);
  if (!cells) return p;
  const s = clamp(start, 0, cells.length);
  const e = clamp(end, s, cells.length);
  const insertCells: CharCell[] = [];
  for (const r of insert) for (const ch of r.text) insertCells.push({ ch, rPr: r.rPr, track: r.track });
  return { ...p, children: runsFromCells([...cells.slice(0, s), ...insertCells, ...cells.slice(e)]) };
}

/** True iff every character in [start, end) has the boolean run prop `key` set on. */
export function rangeUniform(p: Paragraph, start: number, end: number, key: "bold" | "italic" | "strike"): boolean {
  const cells = flattenChars(p);
  if (!cells) return false;
  const s = clamp(start, 0, cells.length);
  const e = clamp(end, s, cells.length);
  if (s === e) return false;
  for (let i = s; i < e; i++) if (cells[i].rPr[key] !== true) return false;
  return true;
}

/** True iff every character in [start, end) is underlined (not "none"). */
export function rangeUnderlined(p: Paragraph, start: number, end: number): boolean {
  const cells = flattenChars(p);
  if (!cells) return false;
  const s = clamp(start, 0, cells.length);
  const e = clamp(end, s, cells.length);
  if (s === e) return false;
  for (let i = s; i < e; i++) {
    const u = cells[i].rPr.underline;
    if (u === undefined || u === "none") return false;
  }
  return true;
}

// ── block operations (split / merge / indent) ────────────────────────────────────────────────────

/**
 * Split a run-only paragraph at character `offset` into two paragraphs that share its pPr (Enter).
 * The run covering the offset is split; formatting is preserved on both sides.
 */
export function splitParagraph(p: Paragraph, offset: number): [Paragraph, Paragraph] {
  const cells = flattenChars(p);
  if (!cells) return [p, { ...p, children: [{ type: "run", rPr: {}, text: "" }] }];
  const at = clamp(offset, 0, cells.length);
  return [
    { ...p, children: runsFromCells(cells.slice(0, at)) },
    { ...p, children: runsFromCells(cells.slice(at)) },
  ];
}

/** Merge paragraph `b` into the end of `a` (Backspace at start of b). Keeps `a`'s pPr. */
export function mergeParagraphs(a: Paragraph, b: Paragraph): Paragraph {
  return { ...a, children: [...a.children, ...b.children] };
}

/** Indent/outdent: change a list paragraph's level, or a normal paragraph's left indent, by `delta`. */
export function setListLevel(p: Paragraph, delta: number): Paragraph {
  if (p.pPr.numbering) {
    const level = clamp(p.pPr.numbering.level + delta, 0, 8);
    return { ...p, pPr: { ...p.pPr, numbering: { ...p.pPr.numbering, level } } };
  }
  const left = Math.max(0, (p.pPr.indent?.left ?? 0) + delta * 720);
  return { ...p, pPr: { ...p.pPr, indent: { ...p.pPr.indent, left } } };
}

// ── track changes (ARCHITECTURE.md §3.12) ────────────────────────────────────────────────────────

/** Map/filter inline nodes (descending into hyperlinks); return null from `fn` to drop a node. */
function mapFilterInlines(nodes: Inline[], fn: (n: Inline) => Inline | null): Inline[] {
  const out: Inline[] = [];
  for (const n of nodes) {
    if (n.type === "hyperlink") { out.push({ ...n, children: mapFilterInlines(n.children, fn) }); continue; }
    const r = fn(n);
    if (r) out.push(r);
  }
  return out;
}

function untrack(r: Run): Run {
  const { track: _drop, ...rest } = r;
  return rest;
}

/** Accept all tracked changes in the paragraph: keep insertions (as plain runs), drop deletions. */
export function acceptChanges(p: Paragraph): Paragraph {
  return {
    ...p,
    children: mapFilterInlines(p.children, (n) =>
      n.type === "run" && n.track ? (n.track.type === "del" ? null : untrack(n)) : n,
    ),
  };
}

/** Reject all tracked changes in the paragraph: drop insertions, restore deletions (as plain runs). */
export function rejectChanges(p: Paragraph): Paragraph {
  return {
    ...p,
    children: mapFilterInlines(p.children, (n) =>
      n.type === "run" && n.track ? (n.track.type === "ins" ? null : untrack(n)) : n,
    ),
  };
}

/** Mark every run as a tracked insertion/deletion (change-mode editing). */
export function markTracked(p: Paragraph, change: TrackChange): Paragraph {
  return mapRuns(p, (r) => ({ ...r, track: change }));
}

/**
 * Convenience: apply a paragraph transform and splice the result back into document.xml at the
 * paragraph's source span. `paragraph` must come from the same parse of `documentXml` (its source
 * offsets must be valid). Returns the new document.xml.
 */
export function applyEdit(documentXml: string, paragraph: Paragraph, transform: (p: Paragraph) => Paragraph): string {
  return patchParagraph(documentXml, transform(paragraph));
}
