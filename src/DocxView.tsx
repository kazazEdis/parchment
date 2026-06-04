// Tier-1 render of a parsed .docx to the screen (ARCHITECTURE.md §3.8). Walks the document model,
// applies the full style cascade + list markers via resolve.ts, maps to CSS via cssMap.ts, and lays
// tables out from the resolved span grid. This is a continuous "page sheet" (no computed pagination
// yet — that is the Tier-2 layout-engine); it is the first thing you can actually look at, and the
// surface the editing layer will mount onto.
import React, { useMemo } from "react";
import { type DocxPackage, getPartText } from "./opc";
import { parseDocument, type Block, type Paragraph, type Inline, type Table } from "./model";
import { parseStyles, type StyleSheet } from "./styles";
import { parseNumbering, type Numbering } from "./numbering";
import { effectiveRunProps, effectiveParagraphProps, markerRunProps, assignListNumbers } from "./resolve";
import { runCss, paragraphCss, drawingCss, trackCss } from "./cssMap";
import { resolveTableGrid } from "./table";
import { resolveImageDataUrl, relationshipTarget } from "./images";
import { ommlToMathML } from "./math";
import { twipsToPx } from "./units";
import type { ParagraphProps } from "./props";

interface Ctx {
  sheet: StyleSheet;
  numbering: Numbering;
  markers: Map<Paragraph, string>;
  pkg: DocxPackage;
}

function renderText(text: string): React.ReactNode[] {
  // Tabs → a fixed gap; explicit line breaks → <br/>.
  const out: React.ReactNode[] = [];
  text.split("\n").forEach((line, li) => {
    if (li > 0) out.push(<br key={`br${li}`} />);
    line.split("\t").forEach((seg, si) => {
      if (si > 0) out.push(<span key={`t${li}-${si}`} style={{ display: "inline-block", width: "0.5in" }} />);
      if (seg) out.push(<React.Fragment key={`s${li}-${si}`}>{seg}</React.Fragment>);
    });
  });
  return out;
}

function InlineView({ node, paraPPr, ctx }: { node: Inline; paraPPr: ParagraphProps; ctx: Ctx }): React.ReactElement | null {
  if (node.type === "run") {
    const css = runCss(effectiveRunProps(ctx.sheet, ctx.numbering, paraPPr, node.rPr));
    return <span style={node.track ? { ...css, ...trackCss(node.track.type) } : css}>{renderText(node.text)}</span>;
  }
  if (node.type === "hyperlink") {
    const href = node.rId ? relationshipTarget(ctx.pkg, node.rId) : node.anchor ? `#${node.anchor}` : undefined;
    return (
      <a href={href} style={{ color: "#0563C1", textDecoration: "underline" }}>
        {node.children.map((c, i) => (
          <InlineView key={i} node={c} paraPPr={paraPPr} ctx={ctx} />
        ))}
      </a>
    );
  }
  if (node.type === "footnoteRef") return <sup style={{ color: "#1C5742", fontSize: "0.7em" }}>{node.id}</sup>;
  if (node.type === "math") return <span dangerouslySetInnerHTML={{ __html: ommlToMathML(node.omml) }} />;
  const src = node.rEmbed ? resolveImageDataUrl(ctx.pkg, node.rEmbed) : undefined;
  if (!src) return null;
  return <img src={src} alt={node.alt ?? ""} style={drawingCss(node.widthEmu, node.heightEmu)} />;
}

function ParagraphView({ p, ctx }: { p: Paragraph; ctx: Ctx }): React.ReactElement {
  const eff = effectiveParagraphProps(ctx.sheet, ctx.numbering, p.pPr);
  const marker = ctx.markers.get(p);
  return (
    <p style={{ margin: 0, ...paragraphCss(eff) }}>
      {marker !== undefined && (
        <span style={{ ...runCss(markerRunProps(ctx.sheet, ctx.numbering, p.pPr)), whiteSpace: "nowrap", marginRight: "0.4em" }}>
          {marker}
        </span>
      )}
      {p.children.length ? p.children.map((n, i) => <InlineView key={i} node={n} paraPPr={p.pPr} ctx={ctx} />) : <br />}
    </p>
  );
}

function TableView({ t, ctx }: { t: Table; ctx: Ctx }): React.ReactElement {
  const grid = resolveTableGrid(t);
  const totalTwips = t.grid.reduce((a, b) => a + b, 0);
  return (
    <table style={{ borderCollapse: "collapse", tableLayout: "fixed", width: totalTwips ? twipsToPx(totalTwips) : "100%", margin: "0 0 8px" }}>
      {totalTwips > 0 && (
        <colgroup>
          {t.grid.map((w, i) => (
            <col key={i} style={{ width: twipsToPx(w) }} />
          ))}
        </colgroup>
      )}
      <tbody>
        {grid.map((cells, ri) => (
          <tr key={ri}>
            {cells.map((rc, ci) => (
              <td
                key={ci}
                colSpan={rc.colSpan > 1 ? rc.colSpan : undefined}
                rowSpan={rc.rowSpan > 1 ? rc.rowSpan : undefined}
                style={{ border: "1px solid #b3b3b3", padding: "2px 5px", verticalAlign: "top" }}
              >
                {rc.cell.blocks.map((b, bi) => (
                  <BlockView key={bi} block={b} ctx={ctx} />
                ))}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function BlockView({ block, ctx }: { block: Block; ctx: Ctx }): React.ReactElement {
  return block.type === "paragraph" ? <ParagraphView p={block} ctx={ctx} /> : <TableView t={block} ctx={ctx} />;
}

/** Render a .docx package as a Word-like page sheet. */
export function DocxView({ pkg, className }: { pkg: DocxPackage; className?: string }): React.ReactElement {
  const { body, page, ctx } = useMemo(() => {
    const model = parseDocument(getPartText(pkg, "word/document.xml") ?? "");
    const sheet = parseStyles(getPartText(pkg, "word/styles.xml") ?? "");
    const numbering = parseNumbering(getPartText(pkg, "word/numbering.xml") ?? "");
    const markers = assignListNumbers(model, numbering);
    const sec = model.section;
    const m = sec?.margins ?? {};
    const page = {
      width: twipsToPx(sec?.pageSize?.width ?? 11906),
      minHeight: twipsToPx(sec?.pageSize?.height ?? 16838),
      paddingTop: twipsToPx(m.top ?? 1440),
      paddingRight: twipsToPx(m.right ?? 1440),
      paddingBottom: twipsToPx(m.bottom ?? 1440),
      paddingLeft: twipsToPx(m.left ?? 1440),
    };
    return { body: model.body, page, ctx: { sheet, numbering, markers, pkg } as Ctx };
  }, [pkg]);

  return (
    <div
      className={className}
      style={{
        width: page.width,
        minHeight: page.minHeight,
        margin: "0 auto",
        background: "#fff",
        boxShadow: "0 1px 6px rgba(0,0,0,0.15)",
        paddingTop: page.paddingTop,
        paddingRight: page.paddingRight,
        paddingBottom: page.paddingBottom,
        paddingLeft: page.paddingLeft,
        boxSizing: "border-box",
        fontFamily: "Calibri, Carlito, sans-serif",
        fontSize: "11pt",
        lineHeight: 1.15,
        color: "#000",
      }}
    >
      {body.map((b, i) => (
        <BlockView key={i} block={b} ctx={ctx} />
      ))}
    </div>
  );
}
