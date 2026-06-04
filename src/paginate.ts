// Computed pagination (ARCHITECTURE.md §3.8, Tier-2). Pack measured block heights into pages of a
// fixed content height — the layout-engine's core, pure + testable. v1 breaks at block boundaries
// (a paragraph/table stays whole on one page); a block taller than a page gets its own page and
// overflows. Line-level splitting within a paragraph is a later refinement.

/**
 * Greedily assign block indices to pages. A page holds as many consecutive blocks as fit in
 * `contentHeight`; a block that doesn't fit starts a new page (and stands alone if taller than a
 * page). Always returns at least one (possibly empty) page.
 */
/**
 * Line-level page breaks (the SuperDoc `sliceLines` idea, studied from its layout-engine). Given the
 * cumulative bottom-Y of every line in the continuous flow (sorted ascending), return the Y offset at
 * which each page begins. Breaks land on line boundaries, so a long paragraph splits across pages
 * without cutting a line in half. A single line taller than a page gets its own page (force-progress).
 * Render each page as a fixed `contentHeight` window of the flow translated by `-breaks[i]`.
 */
export function computePageBreaks(lineBottoms: number[], contentHeight: number): number[] {
  if (lineBottoms.length === 0 || contentHeight <= 0) return [0];
  const breaks = [0];
  let start = 0;
  let i = 0;
  while (i < lineBottoms.length) {
    const limit = start + contentHeight;
    let last = i;
    while (last < lineBottoms.length && lineBottoms[last] <= limit) last++;
    if (last === i) last = i + 1; // oversized line → force one line onto the page
    if (last >= lineBottoms.length) break; // remaining lines all fit on the last page
    const next = lineBottoms[last - 1];
    breaks.push(next);
    start = next;
    i = last;
  }
  return breaks;
}

export function packPages(heights: number[], contentHeight: number): number[][] {
  const pages: number[][] = [];
  let cur: number[] = [];
  let used = 0;
  for (let i = 0; i < heights.length; i++) {
    const h = heights[i];
    if (cur.length > 0 && used + h > contentHeight) {
      pages.push(cur);
      cur = [];
      used = 0;
    }
    cur.push(i);
    used += h;
  }
  if (cur.length) pages.push(cur);
  return pages.length ? pages : [[]];
}
