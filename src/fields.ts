// OOXML fields (ARCHITECTURE.md §3.11). Parse a field instruction (PAGE, NUMPAGES, TOC, REF,
// HYPERLINK, …) and generate a table of contents from heading styles. Live recomputation (PAGE after
// pagination, dynamic TOC) builds on these pure primitives; cached field results already render as
// text once fldSimple is unwrapped in the model.
import { type DocumentModel, type Paragraph, paragraphText } from "./model";

export interface FieldInstruction {
  /** Upper-cased field type, e.g. "PAGE", "TOC", "REF". */
  type: string;
  /** Positional arguments (quotes stripped). */
  args: string[];
  /** Switches: `\o "1-3"` → { "\\o": "1-3" }; a flag switch → true. */
  switches: Record<string, string | true>;
}

/** Parse a field instruction string (the text inside w:instrText / w:fldSimple@w:instr). */
export function parseFieldInstruction(instr: string): FieldInstruction {
  const tokens = instr.trim().match(/"[^"]*"|\S+/g) ?? [];
  const unquote = (s: string): string => s.replace(/^"|"$/g, "");
  const type = (tokens[0] ?? "").toUpperCase();
  const args: string[] = [];
  const switches: Record<string, string | true> = {};
  for (let i = 1; i < tokens.length; i++) {
    const tok = tokens[i];
    if (tok.startsWith("\\")) {
      const next = tokens[i + 1];
      if (next && !next.startsWith("\\")) { switches[tok] = unquote(next); i++; }
      else switches[tok] = true;
    } else {
      args.push(unquote(tok));
    }
  }
  return { type, args, switches };
}

export interface TocEntry {
  level: number;
  text: string;
  styleId: string;
}

/** Build TOC entries from heading-styled paragraphs (styleId "Heading1".."HeadingN" or outlineLevel). */
export function generateToc(model: DocumentModel, maxLevel = 3): TocEntry[] {
  const out: TocEntry[] = [];
  for (const b of model.body) {
    if (b.type !== "paragraph") continue;
    const p: Paragraph = b;
    const styleId = p.pPr.styleId ?? "";
    const m = /^Heading([1-9])$/i.exec(styleId);
    const level = m ? parseInt(m[1], 10) : p.pPr.outlineLevel !== undefined ? p.pPr.outlineLevel + 1 : 0;
    if (level >= 1 && level <= maxLevel) {
      const text = paragraphText(p).trim();
      if (text) out.push({ level, text, styleId });
    }
  }
  return out;
}

/** Resolve a PAGE/NUMPAGES field to its value given the current pagination. */
export function computePageField(type: string, pageNumber: number, totalPages: number): string {
  if (type === "PAGE") return String(pageNumber);
  if (type === "NUMPAGES") return String(totalPages);
  return "";
}
