"use client";

// Computed-pagination viewer (ARCHITECTURE.md §3.8, Tier-2). Renders the document into real page
// sheets: it measures each top-level block's height in an offscreen pass, packs blocks into pages of
// the section's content height (paginate.packPages), then lays out one fixed-size page per group with
// a page-number footer. Block-level breaks in v1 (a paragraph/table stays whole); line-level
// splitting is a later refinement. This is the layout-engine SuperDoc has — here, light and owned.
import React, { useLayoutEffect, useMemo, useRef, useState } from "react";
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

interface Ctx {
  sheet: StyleSheet;
  numbering: Numbering;
  markers: Map<Paragraph, string>;
  pkg: DocxPackage;
  relsPart: string;   // rels file images/links in this container resolve against (body vs header/footer)
  page?: { num: number; total: number };   // current/total page, for live PAGE/NUMPAGES field runs (footer)
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
    const css = runCss(effectiveRunProps(ctx.sheet, ctx.numbering, paraPPr, node.rPr));
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
  const img = <img src={src} alt={node.alt ?? ""} style={drawingCss(node.widthEmu, node.heightEmu)} />;
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
    const eff = effectiveParagraphProps(ctx.sheet, ctx.numbering, b.pPr);
    const marker = ctx.markers.get(b);
    return (
      <p key={key} style={{ margin: 0, ...paragraphCss(eff), ...(hasAnchoredDrawing(b) ? { position: "relative" } : null) }}>
        {marker !== undefined && <span style={{ ...runCss(markerRunProps(ctx.sheet, ctx.numbering, b.pPr)), marginRight: "0.4em" }}>{marker}</span>}
        {b.children.length ? b.children.map((n, i) => <Inline key={i} node={n} paraPPr={b.pPr} ctx={ctx} />) : <br />}
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
  return (
    <table key={key} style={{ borderCollapse: "collapse", tableLayout: "fixed", width: total ? twipsToPx(total) : "100%", margin: "0 0 8px" }}>
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
                  style={{ borderTop: top, borderRight: right, borderBottom: bottom, borderLeft: left, background: rc.cell.props.shd ? `#${rc.cell.props.shd}` : undefined, padding: "2px 5px", verticalAlign: "top" }}>
                  {rc.cell.blocks.map((cb, bi) => renderBlock(cb, ctx, bi))}
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
  const { blocks, header, footer, ctx, headerCtx, footerCtx, geo } = useMemo(() => {
    const model = parseDocument(getPartText(pkg, "word/document.xml") ?? "");
    const sheet = parseStyles(getPartText(pkg, "word/styles.xml") ?? "");
    const numbering = parseNumbering(getPartText(pkg, "word/numbering.xml") ?? "");
    const markers = assignListNumbers(model, numbering);
    const sec = model.section;
    const m = sec?.margins ?? {};
    const padTop = twipsToPx(m.top ?? 1440), padRight = twipsToPx(m.right ?? 1440), padBottom = twipsToPx(m.bottom ?? 1440), padLeft = twipsToPx(m.left ?? 1440);
    const pageWidth = twipsToPx(sec?.pageSize?.width ?? 11906);
    const pageHeight = twipsToPx(sec?.pageSize?.height ?? 16838);
    const footerReserve = 28;
    const geo = { pageWidth, pageHeight, padTop, padRight, padBottom, padLeft, contentWidth: pageWidth - padLeft - padRight, contentHeight: pageHeight - padTop - padBottom - footerReserve };
    const hx = headerXml(pkg, model.section);
    const fx = footerXml(pkg, model.section);
    const header = parseContainer(hx?.xml ?? "", "w:hdr");
    const footer = parseContainer(fx?.xml ?? "", "w:ftr");
    const ctx: Ctx = { sheet, numbering, markers, pkg, relsPart: "word/_rels/document.xml.rels" };
    // Header/footer images resolve against THEIR part's rels (header1.xml.rels), not the body's.
    const headerCtx: Ctx = { ...ctx, relsPart: hx?.relsPart ?? ctx.relsPart };
    const footerCtx: Ctx = { ...ctx, relsPart: fx?.relsPart ?? ctx.relsPart };
    return { blocks: model.body, header, footer, ctx, headerCtx, footerCtx, geo };
  }, [pkg]);

  const blockNodes = useMemo(() => blocks.map((b, i) => renderBlock(b, ctx, i)), [blocks, ctx]);
  const headerNodes = useMemo(() => header.map((b, i) => renderBlock(b, headerCtx, 10000 + i)), [header, headerCtx]);
  const measureRef = useRef<HTMLDivElement | null>(null);
  const [layout, setLayout] = useState<{ breaks: number[]; total: number } | null>(null);

  // Measure every line's bottom-Y in the continuous flow (line-level breaks, the SuperDoc sliceLines
  // idea); break on line boundaries so paragraphs split across pages without cutting a line.
  useLayoutEffect(() => {
    const root = measureRef.current;
    if (!root) return;
    const top = root.getBoundingClientRect().top;
    const bottoms = new Set<number>();
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const range = document.createRange();
      range.selectNodeContents(n);
      for (const r of Array.from(range.getClientRects())) bottoms.add(Math.round(r.bottom - top));
    }
    root.querySelectorAll(":scope > *").forEach((el) => bottoms.add(Math.round(el.getBoundingClientRect().bottom - top)));
    const sorted = [...bottoms].filter((b) => b > 0).sort((a, b) => a - b);
    setLayout({ breaks: computePageBreaks(sorted, geo.contentHeight), total: root.scrollHeight });
  }, [blockNodes, geo.contentWidth, geo.contentHeight]);

  return (
    <div style={{ background: "#e9e6df", padding: 24 }}>
      {/* offscreen measurer (continuous flow at content width) */}
      <div ref={measureRef} style={{ position: "absolute", visibility: "hidden", left: -99999, top: 0, width: geo.contentWidth, fontFamily: "Calibri, Carlito, sans-serif", fontSize: "11pt", lineHeight: 1.15 }}>
        {blockNodes}
      </div>
      {layout?.breaks.map((startY, pi) => {
        const endY = pi < layout.breaks.length - 1 ? layout.breaks[pi + 1] : layout.total;
        return (
          <div
            key={pi}
            className="docx-page"
            style={{ width: geo.pageWidth, height: geo.pageHeight, margin: "0 auto 18px", background: "#fff", boxShadow: "0 1px 6px rgba(0,0,0,0.15)", boxSizing: "border-box", position: "relative", paddingTop: geo.padTop, paddingRight: geo.padRight, paddingBottom: geo.padBottom, paddingLeft: geo.padLeft, fontFamily: "Calibri, Carlito, sans-serif", fontSize: "11pt", lineHeight: 1.15, color: "#000", overflow: "hidden" }}
          >
            {headerNodes.length > 0 ? (
              <div style={{ position: "absolute", top: Math.max(8, geo.padTop / 3), left: geo.padLeft, right: geo.padRight, fontSize: "10pt", color: "#555" }}>{headerNodes}</div>
            ) : headerOverride ? (
              <div style={{ position: "absolute", top: Math.max(8, geo.padTop / 3), left: geo.padLeft, right: geo.padRight }}>{headerOverride}</div>
            ) : null}
            {/* window clipped to this page's line slice; the same flow translated up by startY */}
            <div style={{ height: Math.min(endY - startY, geo.contentHeight), overflow: "hidden" }}>
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
