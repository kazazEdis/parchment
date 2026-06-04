// Footnotes (ARCHITECTURE.md §3.14-ish). Parse word/footnotes.xml and locate w:footnoteReference
// markers in the body. The separator/continuationSeparator pseudo-footnotes are skipped.
import { findElements, getAttr, unescapeXml } from "./xml";

export interface Footnote {
  id: string;
  text: string;
}

export function parseFootnotes(footnotesXml: string | undefined): Footnote[] {
  if (!footnotesXml) return [];
  const out: Footnote[] = [];
  for (const f of findElements(footnotesXml, "w:footnote")) {
    const type = getAttr(f.openTag, "w:type");
    if (type === "separator" || type === "continuationSeparator") continue;
    const inner = footnotesXml.slice(f.innerStart, f.innerEnd);
    const text = findElements(inner, "w:t").map((t) => unescapeXml(inner.slice(t.innerStart, t.innerEnd))).join("");
    out.push({ id: getAttr(f.openTag, "w:id") ?? "", text });
  }
  return out;
}

/** Footnote ids referenced in the body, in document order. */
export function footnoteRefs(documentXml: string): string[] {
  return findElements(documentXml, "w:footnoteReference").map((r) => getAttr(r.openTag, "w:id") ?? "").filter(Boolean);
}
