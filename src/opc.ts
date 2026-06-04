// OPC (Open Packaging Conventions) layer for the in-house .docx editor.
//
// A .docx is a ZIP of XML + media "parts" (ECMA-376 / OOXML). This module reads the zip into an
// in-memory part map and writes it back. The whole fidelity strategy ("preserve-and-patch") rests
// here: we keep EVERY original part's bytes verbatim, so anything the editor doesn't model survives
// a round-trip untouched. The serializer (serialize.ts) only rewrites the parts the user edited.
//
// MIT/permissive only: jszip (already a dependency). No SuperDoc code — this implements the open
// OPC/ZIP container, nothing proprietary.
import JSZip from "jszip";
import { isEncryptedOfficeFile } from "./encrypted";

export const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/** A single package part (one zip entry), content kept as raw bytes for verbatim preservation. */
export interface DocxPart {
  /** Zip path, e.g. "word/document.xml", "[Content_Types].xml", "_rels/.rels". */
  path: string;
  /** Decompressed content. Preserved byte-for-byte unless the editor rewrites this part. */
  bytes: Uint8Array;
  /** True for directory entries (rare in .docx, but preserved if present). */
  dir: boolean;
  /** Original entry timestamp, preserved so re-zips don't gratuitously churn. */
  date?: Date;
  /** True once the editor has modified this part (drives clean/dirty serialization). */
  dirty?: boolean;
}

/** An opened .docx: every part, plus the original entry order (preserved on write). */
export interface DocxPackage {
  parts: Map<string, DocxPart>;
  order: string[];
}

type DocxInput = ArrayBuffer | Uint8Array | Blob | Buffer;

async function toUint8Array(input: DocxInput): Promise<Uint8Array> {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  // Blob (browser + Node 18+) / File
  if (typeof Blob !== "undefined" && input instanceof Blob) {
    return new Uint8Array(await input.arrayBuffer());
  }
  // Node Buffer is a Uint8Array subclass; the first check usually catches it, this is a fallback.
  return new Uint8Array(input as unknown as ArrayBuffer);
}

/**
 * Read a .docx into a part map. Every entry is decompressed into raw bytes and kept; nothing is
 * parsed or discarded here — parsing happens lazily, downstream, only for the parts we model.
 */
export async function readDocx(input: DocxInput): Promise<DocxPackage> {
  const data = await toUint8Array(input);
  if (isEncryptedOfficeFile(data)) {
    throw new Error("This .docx is password-encrypted (agile encryption); decrypt it before opening.");
  }
  const zip = await JSZip.loadAsync(data);

  const parts = new Map<string, DocxPart>();
  const order: string[] = [];

  // JSZip preserves entry order in `zip.files`; capture it so the rewrite matches the original.
  const entries = Object.keys(zip.files);
  for (const path of entries) {
    const entry = zip.files[path];
    order.push(path);
    parts.set(path, {
      path,
      dir: entry.dir,
      date: entry.date ?? undefined,
      bytes: entry.dir ? new Uint8Array(0) : await entry.async("uint8array"),
    });
  }

  return { parts, order };
}

/**
 * Re-zip the package back into .docx bytes, preserving entry order and timestamps. Unmodified parts
 * are written from their preserved bytes — that is what keeps fidelity. Returns raw bytes so this
 * works identically in Node (tests) and the browser; use {@link docxToBlob} for a downloadable Blob.
 */
export async function writeDocx(pkg: DocxPackage): Promise<Uint8Array> {
  const zip = new JSZip();
  for (const path of pkg.order) {
    const part = pkg.parts.get(path);
    if (!part) continue;
    if (part.dir) {
      zip.folder(path.replace(/\/$/, ""));
      continue;
    }
    zip.file(path, part.bytes, { date: part.date });
  }
  // DEFLATE matches how Word/most tooling stores .docx; content (decompressed) is unchanged either way.
  return zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    mimeType: DOCX_MIME,
  });
}

/** Wrap .docx bytes in a Blob (browser download / upload). */
export function docxToBlob(bytes: Uint8Array): Blob {
  return new Blob([bytes as BlobPart], { type: DOCX_MIME });
}

const utf8Decoder = new TextDecoder("utf-8");
const utf8Encoder = new TextEncoder();

export function hasPart(pkg: DocxPackage, path: string): boolean {
  return pkg.parts.has(path);
}

export function getPart(pkg: DocxPackage, path: string): DocxPart | undefined {
  return pkg.parts.get(path);
}

/** Decode a part as UTF-8 text (XML parts). Returns undefined if the part is absent. */
export function getPartText(pkg: DocxPackage, path: string): string | undefined {
  const part = pkg.parts.get(path);
  if (!part || part.dir) return undefined;
  return utf8Decoder.decode(part.bytes);
}

/** Replace a part's content from a UTF-8 string and mark it dirty. Adds the part if new. */
export function setPartText(pkg: DocxPackage, path: string, text: string): void {
  setPartBytes(pkg, path, utf8Encoder.encode(text));
}

/** Replace (or add) a part's raw bytes and mark it dirty. New parts append to the entry order. */
export function setPartBytes(pkg: DocxPackage, path: string, bytes: Uint8Array): void {
  const existing = pkg.parts.get(path);
  if (existing) {
    existing.bytes = bytes;
    existing.dir = false;
    existing.dirty = true;
    return;
  }
  pkg.parts.set(path, { path, bytes, dir: false, dirty: true });
  pkg.order.push(path);
}
