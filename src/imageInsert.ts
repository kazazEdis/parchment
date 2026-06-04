// Insert an image into a .docx: add the media part + relationship + content-type, and produce the
// DrawingML inline run to splice into a paragraph (ARCHITECTURE.md §3.10). Pairs with images.ts
// (which resolves a blip back to a data URL for rendering).
import { type DocxPackage, getPart, setPartBytes } from "./opc";
import { addRelationship, ensureDefaultContentType } from "./opcParts";
import { escapeXmlAttr } from "./xml";

const MIME: Record<string, string> = {
  png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp",
  svg: "image/svg+xml", tiff: "image/tiff", tif: "image/tiff",
};
const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

function nextMediaName(pkg: DocxPackage, ext: string): { path: string; name: string } {
  let n = 1;
  while (getPart(pkg, `word/media/image${n}.${ext}`)) n++;
  return { path: `word/media/image${n}.${ext}`, name: `image${n}.${ext}` };
}

/** The DrawingML run XML for an inline image (size in EMU). */
export function imageRunXml(rId: string, widthEmu: number, heightEmu: number, name: string): string {
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const nm = escapeXmlAttr(name);
  return (
    `<w:r><w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="${WP}">` +
    `<wp:extent cx="${widthEmu}" cy="${heightEmu}"/><wp:docPr id="1" name="${nm}"/>` +
    `<a:graphic xmlns:a="${A}"><a:graphicData uri="${PIC}">` +
    `<pic:pic xmlns:pic="${PIC}">` +
    `<pic:nvPicPr><pic:cNvPr id="1" name="${nm}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rId}" xmlns:r="${R}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
  );
}

/**
 * Register an image in the package (media part + relationship + content-type) and return its
 * relationship id + the inline run XML to splice into a paragraph.
 */
export function insertImage(pkg: DocxPackage, opts: { bytes: Uint8Array; ext: string; widthEmu: number; heightEmu: number; name?: string }): { rId: string; runXml: string } {
  const ext = opts.ext.toLowerCase().replace(/^\./, "");
  const { path, name } = nextMediaName(pkg, ext);
  setPartBytes(pkg, path, opts.bytes);
  ensureDefaultContentType(pkg, ext, MIME[ext] ?? "application/octet-stream");
  const rId = addRelationship(pkg, "word/_rels/document.xml.rels", IMAGE_REL, `media/${name}`);
  return { rId, runXml: imageRunXml(rId, opts.widthEmu, opts.heightEmu, opts.name ?? name) };
}
