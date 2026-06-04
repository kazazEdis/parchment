// Resolve an image relationship id (a:blip/@r:embed) to a data URL for rendering (ARCHITECTURE.md
// §3.10). The blip references a relationship in word/_rels/document.xml.rels whose Target points at
// the media part (e.g. "media/image1.png"); we read those bytes and base64-encode them. Works in the
// browser (btoa) and Node tests (Buffer).
import { type DocxPackage, getPart, getPartText } from "./opc";
import { findElements, getAttr } from "./xml";

const MIME: Record<string, string> = {
  png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp",
  svg: "image/svg+xml", emf: "image/emf", wmf: "image/wmf", tiff: "image/tiff", tif: "image/tiff",
};

/** The Target of a relationship by Id, from word/_rels/document.xml.rels (e.g. for hyperlinks). */
export function relationshipTarget(pkg: DocxPackage, relId: string): string | undefined {
  const rels = getPartText(pkg, "word/_rels/document.xml.rels");
  if (!rels) return undefined;
  for (const rel of findElements(rels, "Relationship")) {
    if (getAttr(rel.openTag, "Id") === relId) return getAttr(rel.openTag, "Target");
  }
  return undefined;
}

function bytesToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("base64");
  let bin = "";
  const CHUNK = 0x8000; // chunk to avoid String.fromCharCode arg-count limits
  for (let i = 0; i < bytes.length; i += CHUNK) {
    bin += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
  }
  return btoa(bin);
}

/** Resolve a blip relationship id to a `data:` URL, or undefined if it can't be found. */
export function resolveImageDataUrl(pkg: DocxPackage, relId: string): string | undefined {
  const target = relationshipTarget(pkg, relId);
  if (!target) return undefined;
  const path = target.startsWith("/") ? target.slice(1) : `word/${target}`;
  const part = getPart(pkg, path);
  if (!part || part.dir) return undefined;
  const ext = (path.split(".").pop() ?? "").toLowerCase();
  return `data:${MIME[ext] ?? "application/octet-stream"};base64,${bytesToBase64(part.bytes)}`;
}
