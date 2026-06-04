// OPC part plumbing: add relationships + content-type registrations so newly-inserted parts (images,
// comments, headers) make a valid package on download. Small, shared by imageInsert / hyperlink /
// comment authoring.
import { type DocxPackage, getPartText, setPartText } from "./opc";
import { findElement, findElements, getAttr, escapeXmlAttr } from "./xml";

const RELS_NS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
const CT_NS = 'xmlns="http://schemas.openxmlformats.org/package/2006/content-types"';

/** Next free rIdN for a .rels part. */
export function nextRelId(relsXml: string | undefined): string {
  let max = 0;
  if (relsXml) {
    for (const r of findElements(relsXml, "Relationship")) {
      const m = /^rId(\d+)$/.exec(getAttr(r.openTag, "Id") ?? "");
      if (m) max = Math.max(max, parseInt(m[1], 10));
    }
  }
  return `rId${max + 1}`;
}

/** Add a Relationship to `relsPath` (creating the part if absent) and return the new id. */
export function addRelationship(pkg: DocxPackage, relsPath: string, type: string, target: string, targetMode?: string): string {
  let rels = getPartText(pkg, relsPath);
  const id = nextRelId(rels);
  const rel = `<Relationship Id="${id}" Type="${escapeXmlAttr(type)}" Target="${escapeXmlAttr(target)}"${targetMode ? ` TargetMode="${escapeXmlAttr(targetMode)}"` : ""}/>`;
  const root = rels ? findElement(rels, "Relationships") : undefined;
  if (!rels || !root) {
    rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships ${RELS_NS}>${rel}</Relationships>`;
  } else {
    rels = rels.slice(0, root.innerEnd) + rel + rels.slice(root.innerEnd);
  }
  setPartText(pkg, relsPath, rels);
  return id;
}

/** Ensure [Content_Types].xml has a Default for `ext` (e.g. png → image/png). */
export function ensureDefaultContentType(pkg: DocxPackage, ext: string, contentType: string): void {
  const ct = getPartText(pkg, "[Content_Types].xml");
  if (!ct) return;
  for (const d of findElements(ct, "Default")) {
    if ((getAttr(d.openTag, "Extension") ?? "").toLowerCase() === ext.toLowerCase()) return;
  }
  const root = findElement(ct, "Types");
  const def = `<Default Extension="${escapeXmlAttr(ext)}" ContentType="${escapeXmlAttr(contentType)}"/>`;
  if (!root) {
    setPartText(pkg, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types ${CT_NS}>${def}</Types>`);
  } else {
    setPartText(pkg, "[Content_Types].xml", ct.slice(0, root.innerStart) + def + ct.slice(root.innerStart));
  }
}

/** Ensure [Content_Types].xml has an Override for a specific part (e.g. comments.xml). */
export function ensureOverrideContentType(pkg: DocxPackage, partName: string, contentType: string): void {
  const ct = getPartText(pkg, "[Content_Types].xml");
  if (!ct) return;
  for (const o of findElements(ct, "Override")) {
    if (getAttr(o.openTag, "PartName") === partName) return;
  }
  const root = findElement(ct, "Types");
  const ov = `<Override PartName="${escapeXmlAttr(partName)}" ContentType="${escapeXmlAttr(contentType)}"/>`;
  if (root) setPartText(pkg, "[Content_Types].xml", ct.slice(0, root.innerEnd) + ov + ct.slice(root.innerEnd));
}
