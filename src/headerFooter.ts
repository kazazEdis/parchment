// Headers & footers (ARCHITECTURE.md §3.8). A section references header/footer parts by relationship
// id (w:headerReference/@r:id); resolve those to their part XML (root w:hdr / w:ftr) so the paginated
// view can render them at the top/bottom of each page. Types: "default" | "first" | "even".
import { type DocxPackage, getPartText } from "./opc";
import { relationshipTarget } from "./images";
import type { SectionProps } from "./model";

function partFor(pkg: DocxPackage, rId: string): string | undefined {
  const target = relationshipTarget(pkg, rId);
  if (!target) return undefined;
  return getPartText(pkg, target.startsWith("/") ? target.slice(1) : `word/${target}`);
}

function pick(refs: { type: string; rId: string }[] | undefined, prefer: string): { type: string; rId: string } | undefined {
  if (!refs || refs.length === 0) return undefined;
  return refs.find((r) => r.type === prefer) ?? refs.find((r) => r.type === "default") ?? refs[0];
}

/** XML of the section's header part (w:hdr root), or undefined. */
export function headerXml(pkg: DocxPackage, section: SectionProps | undefined, prefer: "default" | "first" | "even" = "default"): string | undefined {
  const ref = pick(section?.headerRefs, prefer);
  return ref ? partFor(pkg, ref.rId) : undefined;
}

/** XML of the section's footer part (w:ftr root), or undefined. */
export function footerXml(pkg: DocxPackage, section: SectionProps | undefined, prefer: "default" | "first" | "even" = "default"): string | undefined {
  const ref = pick(section?.footerRefs, prefer);
  return ref ? partFor(pkg, ref.rId) : undefined;
}
