// Headers & footers (ARCHITECTURE.md §3.8). A section references header/footer parts by relationship
// id (w:headerReference/@r:id); resolve those to their part XML (root w:hdr / w:ftr) so the paginated
// view can render them at the top/bottom of each page. Types: "default" | "first" | "even".
import { type DocxPackage, getPartText } from "./opc";
import { relationshipTarget, relsPathFor } from "./images";
import type { SectionProps } from "./model";

// XML of a header/footer part + the rels file that part's images/links resolve against (so the
// renderer can load header/footer logos, whose relationships live in e.g. header1.xml.rels — NOT
// the document body's rels).
export interface PartXml { xml: string; relsPart: string; }

function partFor(pkg: DocxPackage, rId: string): PartXml | undefined {
  const target = relationshipTarget(pkg, rId);
  if (!target) return undefined;
  const partPath = target.startsWith("/") ? target.slice(1) : `word/${target}`;
  const xml = getPartText(pkg, partPath);
  return xml === undefined ? undefined : { xml, relsPart: relsPathFor(partPath) };
}

function pick(refs: { type: string; rId: string }[] | undefined, prefer: string): { type: string; rId: string } | undefined {
  if (!refs || refs.length === 0) return undefined;
  return refs.find((r) => r.type === prefer) ?? refs.find((r) => r.type === "default") ?? refs[0];
}

/** The section's header part (w:hdr root XML + its rels path), or undefined. */
export function headerXml(pkg: DocxPackage, section: SectionProps | undefined, prefer: "default" | "first" | "even" = "default"): PartXml | undefined {
  const ref = pick(section?.headerRefs, prefer);
  return ref ? partFor(pkg, ref.rId) : undefined;
}

/** The section's footer part (w:ftr root XML + its rels path), or undefined. */
export function footerXml(pkg: DocxPackage, section: SectionProps | undefined, prefer: "default" | "first" | "even" = "default"): PartXml | undefined {
  const ref = pick(section?.footerRefs, prefer);
  return ref ? partFor(pkg, ref.rId) : undefined;
}
