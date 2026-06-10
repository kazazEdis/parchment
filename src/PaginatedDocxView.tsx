"use client";

// Computed-pagination viewer (ARCHITECTURE.md §3.8, Tier-2). Renders the document into real page
// sheets: it measures each top-level block's height in an offscreen pass, packs blocks into pages of
// the section's content height (paginate.packPages), then lays out one fixed-size page per group with
// a page-number footer. Block-level breaks in v1 (a paragraph/table stays whole); line-level
// splitting is a later refinement. This is the layout-engine SuperDoc has — here, light and owned.
import React, { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { type DocxPackage, getPartText } from "./opc";
import { parseDocument, parseContainer, type Block, type Paragraph, type Inline, type BorderSide } from "./model";
import { headerXml, footerXml } from "./headerFooter";
import { parseStyles, resolveTableStyleBorders, type StyleSheet } from "./styles";
import { parseNumbering, type Numbering } from "./numbering";
import { effectiveParagraphProps, effectiveRunProps, markerRunProps, assignListNumbers } from "./resolve";
import { paragraphCss, runCss, trackCss, drawingCss } from "./cssMap";
import { resolveTableGrid } from "./table";
import { resolveImageDataUrl, relationshipTarget } from "./images";
import { ommlToMathML } from "./math";
import { twipsToPx, emuToPx } from "./units";
import { computePageBreaks } from "./paginate";
import { ensureFontsInjected } from "./fonts";

interface Ctx {
  sheet: StyleSheet;
  numbering: Numbering;
  markers: Map<Paragraph, string>;
  pkg: DocxPackage;
  relsPart: string;   // rels file images/links in this container resolve against (body vs header/footer)
  page?: { num: number; total: number };   // current/total page, for live PAGE/NUMPAGES field runs (footer)
  tableStyleId?: string;   // style of the enclosing table → applies to cell paragraphs/runs (Word cascade)
}

const renderText = (text: string): React.ReactNode[] => {
  const out: React.ReactNode[] = [];
  text.split("\n").forEach((line, li) => {
    if (li > 0) out.push(<br key={`b${li}`} />);
    line.split("\t").forEach((seg, si) => {
      if (si > 0) out.push(<span key={`t${li}-${si}`} style={{ display: "inline-block", width: "0.5in" }} />);
      if (seg) out.push(<React.Fragment key={`s${li}-${si}`}>{seg}</React.Fragment>);
    });
  });
  return out;
};

function Inline({ node, paraPPr, ctx }: { node: Inline; paraPPr: Paragraph["pPr"]; ctx: Ctx }): React.ReactElement | null {
  if (node.type === "run") {
    const css = runCss(effectiveRunProps(ctx.sheet, ctx.numbering, paraPPr, node.rPr, ctx.tableStyleId));
    // Live page-number fields (footer): show the current/total page, not the stale cached text.
    const text = node.field && ctx.page ? String(node.field === "PAGE" ? ctx.page.num : ctx.page.total) : node.text;
    return <span style={node.track ? { ...css, ...trackCss(node.track.type) } : css}>{renderText(text)}</span>;
  }
  if (node.type === "hyperlink") {
    const href = node.rId ? relationshipTarget(ctx.pkg, node.rId, ctx.relsPart) : node.anchor ? `#${node.anchor}` : undefined;
    return <a href={href} style={{ color: "#0563C1", textDecoration: "underline" }}>{node.children.map((c, i) => <Inline key={i} node={c} paraPPr={paraPPr} ctx={ctx} />)}</a>;
  }
  if (node.type === "footnoteRef") return <sup style={{ color: "#1C5742", fontSize: "0.7em" }}>{node.id}</sup>;
  if (node.type === "math") return <span dangerouslySetInnerHTML={{ __html: ommlToMathML(node.omml) }} />;
  const src = node.rEmbed ? resolveImageDataUrl(ctx.pkg, node.rEmbed, ctx.relsPart) : undefined;
  if (!src) return null;
  let img: React.ReactElement;
  if (node.crop) {
    // a:srcRect: show only the visible sub-region scaled into the extent box. Scale the full image so
    // its visible fraction fills the box, clip with an overflow-hidden wrapper, and offset by the crop.
    const { l, t, r, b } = node.crop;
    const vw = Math.max(0.0001, 1 - l - r), vh = Math.max(0.0001, 1 - t - b);
    const w = emuToPx(node.widthEmu ?? 0), h = emuToPx(node.heightEmu ?? 0);
    img = (
      <span style={{ display: "inline-block", width: w, height: h, overflow: "hidden", position: "relative", verticalAlign: "bottom" }}>
        <img src={src} alt={node.alt ?? ""} style={{ position: "absolute", width: w / vw, height: h / vh, left: -(l / vw) * w, top: -(t / vh) * h, maxWidth: "none" }} />
      </span>
    );
  } else {
    img = <img src={src} alt={node.alt ?? ""} style={drawingCss(node.widthEmu, node.heightEmu)} />;
  }
  // Anchored (floating) drawing → out of flow, absolutely positioned at its posOffset relative to the
  // paragraph (which renderBlock marks position:relative). Inline drawings render in flow as before.
  if (node.anchored && (node.anchorXEmu != null || node.anchorYEmu != null)) {
    return (
      <span style={{ position: "absolute", left: emuToPx(node.anchorXEmu ?? 0), top: emuToPx(node.anchorYEmu ?? 0), zIndex: node.behindDoc ? 0 : 2, lineHeight: 0 }}>
        {img}
      </span>
    );
  }
  return img;
}

// Does a paragraph hold an anchored (floating) drawing? Such paragraphs render position:relative so
// the absolutely-positioned image is placed relative to them.
function hasAnchoredDrawing(p: Paragraph): boolean {
  return p.children.some((n) => n.type === "drawing" && n.anchored && (n.anchorXEmu != null || n.anchorYEmu != null));
}

// True when none of a paragraph's children contribute IN-FLOW height — every child is either an
// anchored (floating) drawing (absolutely positioned, no line box) or a whitespace-only run (CSS
// white-space:normal collapses a lone space to nothing → no line box). Word still reserves a line for
// such a paragraph (the paragraph mark), so we emit a <br> to match; without it the paragraph
// collapses to 0px and content anchored to a LATER paragraph (e.g. a header logo's badge row) rides
// up over the earlier floating images.
function isFloatingOnly(p: Paragraph): boolean {
  return p.children.length > 0 && p.children.every(
    (n) => (n.type === "drawing" && n.anchored) || (n.type === "run" && !(n.text ?? "").trim()),
  );
}

// One OOXML border edge → a CSS border value, or undefined for "no edge". val "none"/"nil" (or an
// absent side) means Word draws nothing — the previous code drew a 1px rule on EVERY cell regardless,
// turning borderless tables into a heavy grid. sz is in eighths of a point.
function borderCss(s?: BorderSide): string | undefined {
  if (!s || s.val === "none" || s.val === "nil") return undefined;
  const px = Math.max(1, Math.round((s.sz / 8) * (96 / 72)));
  const color = !s.color || s.color === "auto" ? "#000" : `#${s.color}`;
  return `${px}px solid ${color}`;
}

function renderBlock(b: Block, ctx: Ctx, key: number): React.ReactElement {
  if (b.type === "paragraph") {
    const eff = effectiveParagraphProps(ctx.sheet, ctx.numbering, b.pPr, ctx.tableStyleId);
    const marker = ctx.markers.get(b);
    return (
      <p key={key} style={{ margin: 0, ...paragraphCss(eff), ...(hasAnchoredDrawing(b) ? { position: "relative" } : null) }}>
        {marker !== undefined && <span style={{ ...runCss(markerRunProps(ctx.sheet, ctx.numbering, b.pPr)), marginRight: "0.4em" }}>{marker}</span>}
        {isFloatingOnly(b) ? (
          // Render only the floating drawings (absolute, no line box) + exactly ONE <br> so the
          // paragraph is a single line like Word's paragraph mark. Skip the whitespace runs — they'd
          // either collapse to 0 or, combined with the <br>, render a spurious second line.
          <>
            {b.children.filter((n) => n.type === "drawing").map((n, i) => <Inline key={i} node={n} paraPPr={b.pPr} ctx={ctx} />)}
            <br />
          </>
        ) : b.children.length ? (
          b.children.map((n, i) => <Inline key={i} node={n} paraPPr={b.pPr} ctx={ctx} />)
        ) : (
          <br />
        )}
      </p>
    );
  }
  const total = b.grid.reduce((a, c) => a + c, 0);
  const grid = resolveTableGrid(b);
  // Effective table borders = the table style's borders (e.g. "TableGrid") with the table's own
  // inline w:tblBorders overriding per side. Cell w:tcBorders override again at the cell level below.
  const styleBorders = resolveTableStyleBorders(ctx.sheet, b.styleId);
  const tb: typeof b.borders = styleBorders || b.borders ? { ...styleBorders, ...b.borders } : undefined;
  const lastRow = grid.length - 1;
  const totalCols = b.grid.length;
  // Cell paragraphs resolve against the TABLE STYLE too (e.g. TableGrid sets spacing-after 0 / single
  // line) — Word's cascade. Nested tables override with their own style id.
  const cellCtx: Ctx = b.styleId ? { ...ctx, tableStyleId: b.styleId } : ctx;
  return (
    <table key={key} style={{ borderCollapse: "collapse", tableLayout: "fixed", width: total ? twipsToPx(total) : "100%", margin: "0 0 8px" }}>
      {/* w:tblGrid → <colgroup>: without explicit column widths, table-layout:fixed distributes
          the width EQUALLY across columns, ignoring the docx grid (narrow label + wide content
          columns collapse to uniform stripes and every cell mis-wraps). */}
      {total > 0 && (
        <colgroup>
          {b.grid.map((w, i) => <col key={i} style={{ width: `${((w / total) * 100).toFixed(3)}%` }} />)}
        </colgroup>
      )}
      <tbody>
        {grid.map((cells, ri) => {
          let col = 0;   // running grid column so first/last-column edges resolve to the table's outer border
          return (
            <tr key={ri}>{cells.map((rc, ci) => {
              const cbd = rc.cell.props.borders;   // cell w:tcBorders override the table border per side
              const firstCol = col === 0;
              const lastCol = col + rc.colSpan >= totalCols;
              // A side = the cell's own border if set, else the table's outer edge (first/last) or its
              // interior rule (insideH between rows / insideV between columns).
              const top = borderCss(cbd?.top ?? (ri === 0 ? tb?.top : tb?.insideH));
              const bottom = borderCss(cbd?.bottom ?? (ri === lastRow ? tb?.bottom : tb?.insideH));
              const left = borderCss(cbd?.left ?? (firstCol ? tb?.left : tb?.insideV));
              const right = borderCss(cbd?.right ?? (lastCol ? tb?.right : tb?.insideV));
              col += rc.colSpan;
              return (
                <td key={ci} colSpan={rc.colSpan > 1 ? rc.colSpan : undefined} rowSpan={rc.rowSpan > 1 ? rc.rowSpan : undefined}
                  style={{ borderTop: top, borderRight: right, borderBottom: bottom, borderLeft: left, background: rc.cell.props.shd ? `#${rc.cell.props.shd}` : undefined, padding: "2px 5px", verticalAlign: rc.cell.props.vAlign === "center" ? "middle" : (rc.cell.props.vAlign ?? "top") }}>
                  {rc.cell.blocks.map((cb, bi) => renderBlock(cb, cellCtx, bi))}
                </td>
              );
            })}</tr>
          );
        })}
      </tbody>
    </table>
  );
}

export function PaginatedDocxView({ pkg, headerOverride, footerOverride }: {
  pkg: DocxPackage;
  /** Rendered at the top/bottom of every page ONLY when the docx has no header/footer of its own
   *  (e.g. a host app's letterhead fallback). When the docx supplies one, it wins. */
  headerOverride?: React.ReactNode;
  footerOverride?: React.ReactNode;
}): React.ReactElement {
  const { blocks, headerDefault, headerFirst, footer, hasTitlePage, ctx, headerCtxDefault, headerCtxFirst, footerCtx, geo } = useMemo(() => {
    const model = parseDocument(getPartText(pkg, "word/document.xml") ?? "");
    const sheet = parseStyles(getPartText(pkg, "word/styles.xml") ?? "", getPartText(pkg, "word/theme/theme1.xml"));
    const numbering = parseNumbering(getPartText(pkg, "word/numbering.xml") ?? "");
    const markers = assignListNumbers(model, numbering);
    const sec = model.section;
    const m = sec?.margins ?? {};
    const padTop = twipsToPx(m.top ?? 1440), padRight = twipsToPx(m.right ?? 1440), padBottom = twipsToPx(m.bottom ?? 1440), padLeft = twipsToPx(m.left ?? 1440);
    const pageWidth = twipsToPx(sec?.pageSize?.width ?? 11906);
    const pageHeight = twipsToPx(sec?.pageSize?.height ?? 16838);
    const footerReserve = 28;
    const headerTop = Math.max(8, padTop / 3);   // where the header band starts (mirrors the render below)
    const geo = { pageWidth, pageHeight, padTop, padRight, padBottom, padLeft, headerTop, contentWidth: pageWidth - padLeft - padRight, contentHeight: pageHeight - padTop - padBottom - footerReserve };
    // Default header + (when the section has a title page) the distinct FIRST-page header — e.g. a
    // logo letterhead shown only on page 1. parseContainer("") → [] so an absent ref is harmless.
    const hxDefault = headerXml(pkg, model.section, "default");
    const hxFirst = sec?.titlePage ? headerXml(pkg, model.section, "first") : undefined;
    const fx = footerXml(pkg, model.section);
    const ctx: Ctx = { sheet, numbering, markers, pkg, relsPart: "word/_rels/document.xml.rels" };
    // Header/footer images resolve against THEIR part's rels (header1.xml.rels), not the body's.
    return {
      blocks: model.body,
      headerDefault: parseContainer(hxDefault?.xml ?? "", "w:hdr"),
      headerFirst: hxFirst ? parseContainer(hxFirst.xml, "w:hdr") : [],
      footer: parseContainer(fx?.xml ?? "", "w:ftr"),
      hasTitlePage: !!sec?.titlePage,
      ctx,
      headerCtxDefault: { ...ctx, relsPart: hxDefault?.relsPart ?? ctx.relsPart } as Ctx,
      headerCtxFirst: { ...ctx, relsPart: hxFirst?.relsPart ?? ctx.relsPart } as Ctx,
      footerCtx: { ...ctx, relsPart: fx?.relsPart ?? ctx.relsPart } as Ctx,
      geo,
    };
  }, [pkg]);

  const blockNodes = useMemo(() => blocks.map((b, i) => renderBlock(b, ctx, i)), [blocks, ctx]);
  const headerNodesDefault = useMemo(() => headerDefault.map((b, i) => renderBlock(b, headerCtxDefault, 10000 + i)), [headerDefault, headerCtxDefault]);
  const headerNodesFirst = useMemo(() => headerFirst.map((b, i) => renderBlock(b, headerCtxFirst, 12000 + i)), [headerFirst, headerCtxFirst]);
  // The first-page header (logo band) shown on page 1 only; falls back to the default when absent.
  const usesFirstHeader = hasTitlePage && headerNodesFirst.length > 0;
  const measureRef = useRef<HTMLDivElement | null>(null);
  const firstHeaderRef = useRef<HTMLDivElement | null>(null);
  const defaultHeaderRef = useRef<HTMLDivElement | null>(null);
  const [layout, setLayout] = useState<{ breaks: number[]; total: number; firstExtra: number; defaultExtra: number } | null>(null);
  // Inject the embedded metric-compatible fonts (Calibri≈Carlito) so measurement uses Word's glyph
  // advances. They load asynchronously, so flip `fontsReady` once ready to RE-measure (the first
  // layout pass runs against the fallback font; this corrects the line breaks + pagination).
  const [fontsReady, setFontsReady] = useState(false);
  useEffect(() => {
    ensureFontsInjected();
    let live = true;
    const fonts = (document as any).fonts;
    // EXPLICITLY load Carlito (both weights) before flipping fontsReady → re-measure. @font-face is
    // lazy, so `fonts.ready` can resolve BEFORE Carlito is even requested; measuring then uses the
    // fallback's (shorter) header height → too small a first-page push-down → the body overlaps the
    // header. Forcing the load makes the re-measure see the real Carlito metrics.
    if (fonts?.load) {
      Promise.all([fonts.load("400 12pt Carlito"), fonts.load("700 12pt Carlito")])
        .catch(() => {})
        .then(() => { if (live) setFontsReady(true); });
    } else setFontsReady(true);
    return () => { live = false; };
  }, []);

  // Measure every line's bottom-Y in the continuous flow (line-level breaks, the SuperDoc sliceLines
  // idea); break on line boundaries so paragraphs split across pages without cutting a line. Also
  // measure the first-page header: when it's taller than the top-margin band it reserves extra space,
  // so page 1 gets a shorter content height (firstExtra) and its body is pushed down to clear it.
  useLayoutEffect(() => {
    const root = measureRef.current;
    if (!root) return;
    const rootRect = root.getBoundingClientRect();
    // getBoundingClientRect/getClientRects return POST-transform sizes. A host (e.g. a fit-to-pane
    // preview) may wrap us in transform:scale(), which would shrink every measurement → a too-small
    // first-page push-down (body overlaps the header) and wrong page breaks. The measurer is laid out
    // at geo.contentWidth, so rect.width / contentWidth recovers the applied scale; divide it back out.
    const scale = rootRect.width > 0 ? rootRect.width / geo.contentWidth : 1;
    const top = rootRect.top;
    const bottoms = new Set<number>();
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const range = document.createRange();
      range.selectNodeContents(n);
      for (const r of Array.from(range.getClientRects())) bottoms.add(Math.round((r.bottom - top) / scale));
    }
    root.querySelectorAll(":scope > *").forEach((el) => bottoms.add(Math.round((el.getBoundingClientRect().bottom - top) / scale)));
    const sorted = [...bottoms].filter((b) => b > 0).sort((a, b) => a - b);
    // Table rows must not be sliced across a page boundary (the slice cuts cell text mid-line and the
    // continuation overlaps the row's border on the next page). Collect each row's [top, bottom] flow
    // range as an atomic region so a break lands at the row's top instead of through it.
    const rowRanges: [number, number][] = [];
    root.querySelectorAll("tr").forEach((tr) => {
      const r = tr.getBoundingClientRect();
      rowRanges.push([Math.round((r.top - top) / scale), Math.round((r.bottom - top) / scale)]);
    });
    // How far the first-page header pokes below the normal top margin. The logo band usually fits the
    // top-margin slack, so the page-1 body is only NUDGED down by this (into the same slack) — page 1
    // keeps its full content height, matching Word (which doesn't drop a row to show the letterhead).
    // Clamp so a very tall header can't push the body into the footer.
    // Header height for the body push-down uses the FLOW height only. Word lets floating header
    // images (a logo) sit in the top-margin slack and does NOT push the body down to clear them — so
    // we must not count anchored drawings here, or the body is shoved too far down (the terms table
    // spills to the next page). The logo-vs-rule alignment is handled by correct line metrics, not by
    // pushing the body.
    let firstExtra = 0;
    if (usesFirstHeader && firstHeaderRef.current) {
      const hH = firstHeaderRef.current.getBoundingClientRect().height / scale;
      firstExtra = Math.min(geo.padBottom, Math.max(0, Math.round(geo.headerTop + hH + 6 - geo.padTop)));
    }
    let defaultExtra = 0;
    if (defaultHeaderRef.current) {
      const hH = defaultHeaderRef.current.getBoundingClientRect().height / scale;
      defaultExtra = Math.min(geo.padBottom, Math.max(0, Math.round(geo.headerTop + hH + 6 - geo.padTop)));
    }
    // Later pages reserve the default header's overflow; page 1 reserves the first-page header's
    // overflow when there is one (it shows the first header, not the default), else the default's.
    const body = geo.contentHeight - defaultExtra;
    const firstBody = geo.contentHeight - (usesFirstHeader ? firstExtra : defaultExtra);
    setLayout({ breaks: computePageBreaks(sorted, body, firstBody, rowRanges), total: root.scrollHeight, firstExtra, defaultExtra });
  }, [blockNodes, headerNodesDefault, headerNodesFirst, usesFirstHeader, geo.contentWidth, geo.contentHeight, geo.headerTop, geo.padTop, geo.padBottom, fontsReady]);

  return (
    <div style={{ background: "#e9e6df", padding: 24 }}>
      {/* offscreen measurer (continuous flow at content width) */}
      <div ref={measureRef} style={{ position: "absolute", visibility: "hidden", left: -99999, top: 0, width: geo.contentWidth, fontFamily: "Calibri, Carlito, sans-serif", fontSize: "11pt", lineHeight: 1.15 }}>
        {blockNodes}
      </div>
      {/* offscreen measurer for the first-page header (to reserve its height on page 1) */}
      {usesFirstHeader && (
        <div ref={firstHeaderRef} style={{ position: "absolute", visibility: "hidden", left: -99999, top: 0, width: geo.contentWidth, fontSize: "10pt" }}>
          {headerNodesFirst}
        </div>
      )}
      {/* offscreen measurer for the default header (reserves its height on every page it appears) */}
      {headerNodesDefault.length > 0 && (
        <div ref={defaultHeaderRef} style={{ position: "absolute", visibility: "hidden", left: -99999, top: 0, width: geo.contentWidth, fontSize: "10pt" }}>
          {headerNodesDefault}
        </div>
      )}
      {layout?.breaks.map((startY, pi) => {
        const endY = pi < layout.breaks.length - 1 ? layout.breaks[pi + 1] : layout.total;
        const onFirst = pi === 0 && usesFirstHeader;              // page 1 shows the first-page header
        const pageHeader = onFirst ? headerNodesFirst : headerNodesDefault;
        // Nudge each page's body below its header: the first-page header on page 1, the default
        // header everywhere it appears.
        const bodyExtra = onFirst ? layout.firstExtra : layout.defaultExtra;
        return (
          <div
            key={pi}
            className="docx-page"
            style={{ width: geo.pageWidth, height: geo.pageHeight, margin: "0 auto 18px", background: "#fff", boxShadow: "0 1px 6px rgba(0,0,0,0.15)", boxSizing: "border-box", position: "relative", paddingTop: geo.padTop, paddingRight: geo.padRight, paddingBottom: geo.padBottom, paddingLeft: geo.padLeft, fontFamily: "Calibri, Carlito, sans-serif", fontSize: "11pt", lineHeight: 1.15, color: "#000", overflow: "hidden" }}
          >
            {pageHeader.length > 0 ? (
              <div style={{ position: "absolute", top: geo.headerTop, left: geo.padLeft, right: geo.padRight, fontSize: "10pt", color: "#555" }}>{pageHeader}</div>
            ) : headerOverride ? (
              <div style={{ position: "absolute", top: geo.headerTop, left: geo.padLeft, right: geo.padRight }}>{headerOverride}</div>
            ) : null}
            {/* window clipped to this page's line slice; the same flow translated up by startY. On page 1 a
                tall first-page header pushes the body down by bodyExtra (and shortens its content height). */}
            <div style={{ marginTop: bodyExtra, height: Math.min(endY - startY, geo.contentHeight - bodyExtra), overflow: "hidden" }}>
              <div style={{ transform: `translateY(${-startY}px)`, width: geo.contentWidth }}>{blockNodes}</div>
            </div>
            {footer.length > 0 ? (
              // Render the docx footer PER PAGE so PAGE/NUMPAGES field runs resolve to this page.
              <div style={{ position: "absolute", bottom: Math.max(24, geo.padBottom / 3), left: geo.padLeft, right: geo.padRight, fontSize: "10pt", color: "#555" }}>
                {footer.map((b, i) => renderBlock(b, { ...footerCtx, page: { num: pi + 1, total: layout.breaks.length } }, 20000 + i))}
              </div>
            ) : footerOverride ? (
              <div style={{ position: "absolute", bottom: Math.max(24, geo.padBottom / 3), left: geo.padLeft, right: geo.padRight }}>{footerOverride}</div>
            ) : (
              // No docx footer / override → a minimal page-number stamp so the page is still numbered.
              <div style={{ position: "absolute", bottom: 8, left: 0, right: 0, textAlign: "center", fontSize: "9pt", color: "#777" }}>{pi + 1} / {layout.breaks.length}</div>
            )}
          </div>
        );
      })}
    </div>
  );
}
