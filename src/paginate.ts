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
export function computePageBreaks(
  lineBottoms: number[],
  contentHeight: number,
  firstContentHeight = contentHeight,
  // [top, bottom] flow-Y ranges that must NOT be split across a page (table rows). A break that would
  // land strictly inside one is pulled back to the range's top so the whole row moves to the next page
  // — UNLESS the row is taller than a page or already straddles the page start (then splitting is the
  // only way to make progress). Mirrors Word's default "keep row together where possible".
  atomicRanges: [number, number][] = [],
): number[] {
  if (lineBottoms.length === 0 || contentHeight <= 0) return [0];
  const ranges = [...atomicRanges].sort((a, b) => a[0] - b[0]);
  // Pull a candidate break Y out of any atomic row it bisects, back to that row's top.
  const avoidSplit = (y: number, start: number, pageHeight: number): number => {
    for (const [rTop, rBottom] of ranges) {
      if (rTop < y && y < rBottom) {
        // Only pull back when the whole row would then fit on the next page and pulling makes forward
        // progress (row starts after this page's start). Otherwise allow the split.
        if (rTop > start && rBottom - rTop <= pageHeight) return rTop;
        return y;
      }
    }
    return y;
  };
  const breaks = [0];
  let start = 0;
  let i = 0;
  while (i < lineBottoms.length) {
    // The first page may be shorter (a tall first-page header — e.g. a logo letterhead — reserves
    // vertical space the body must clear). Every later page uses the full content height.
    const pageHeight = breaks.length === 1 ? Math.max(1, firstContentHeight) : contentHeight;
    const limit = start + pageHeight;
    let last = i;
    while (last < lineBottoms.length && lineBottoms[last] <= limit) last++;
    if (last === i) last = i + 1; // oversized line → force one line onto the page
    if (last >= lineBottoms.length) break; // remaining lines all fit on the last page
    let next = avoidSplit(lineBottoms[last - 1], start, pageHeight);
    if (next <= start) next = lineBottoms[last - 1]; // safety: never stall or go backwards
    breaks.push(next);
    start = next;
    // Continue from the first line strictly below the (possibly pulled-back) break.
    i = last;
    while (i > 0 && lineBottoms[i - 1] >= next) i--;
    while (i < lineBottoms.length && lineBottoms[i] <= next) i++;
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
