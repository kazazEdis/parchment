// Word-level diff → tracked-change redlines. Given an original and a revised text, produce runs
// marked w:ins / w:del so the difference renders as redlines in our viewer AND in Word/any
// OOXML-compatible tool. SuperDoc renders *existing* tracked changes; generating tracked changes as a
// diff between two arbitrary versions is a distinctive contract/offer-review capability.
import { type Run, type Paragraph, type TrackChange, paragraphText } from "./model";
import type { RunProps } from "./props";

/** Tokenize into word and whitespace tokens (whitespace kept so spacing survives the diff). */
function tokenize(s: string): string[] {
  return s.match(/\s+|\S+/g) ?? [];
}

type Seg = { type: "eq" | "del" | "ins"; text: string };

/** Longest-common-subsequence diff over token arrays → a coalesced edit script. */
export function diffTokens(a: string[], b: string[]): Seg[] {
  const n = a.length;
  const m = b.length;
  const dp: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }
  const out: Seg[] = [];
  const push = (type: Seg["type"], text: string): void => {
    const last = out[out.length - 1];
    if (last && last.type === type) last.text += text;
    else out.push({ type, text });
  };
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) { push("eq", a[i]); i++; j++; }
    else if (dp[i + 1][j] >= dp[i][j + 1]) { push("del", a[i]); i++; }
    else { push("ins", b[j]); j++; }
  }
  while (i < n) { push("del", a[i++]); }
  while (j < m) { push("ins", b[j++]); }
  return out;
}

export interface RedlineMeta {
  author?: string;
  date?: string;
  id?: string;
}

/** Diff `oldText` → `newText` into runs: unchanged plain, removed as w:del, added as w:ins. */
export function redlineRuns(oldText: string, newText: string, meta: RedlineMeta = {}, rPr: RunProps = {}): Run[] {
  const mk = (text: string, type?: "ins" | "del"): Run => {
    const r: Run = { type: "run", rPr, text };
    if (type) r.track = { type, ...meta } as TrackChange;
    return r;
  };
  return diffTokens(tokenize(oldText), tokenize(newText)).map((seg) =>
    seg.type === "eq" ? mk(seg.text) : seg.type === "del" ? mk(seg.text, "del") : mk(seg.text, "ins"),
  );
}

/**
 * Replace a paragraph's content with the redlined diff of its current text vs `newText`. The diff
 * runs inherit the paragraph's first run formatting. Pair with serialize.patchParagraph to write it.
 */
export function redlineParagraph(p: Paragraph, newText: string, meta: RedlineMeta = {}): Paragraph {
  const firstRun = p.children.find((n): n is Run => n.type === "run");
  return { ...p, children: redlineRuns(paragraphText(p), newText, meta, firstRun ? firstRun.rPr : {}) };
}
