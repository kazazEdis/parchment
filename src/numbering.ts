// numbering.xml — list definitions and marker rendering (ARCHITECTURE.md §3.6).
//
// Lists are not a tree in OOXML: a paragraph just references w:numId + w:ilvl. A w:num points at a
// w:abstractNum (optionally overriding levels), and each abstract level (w:lvl) defines the format
// (decimal/bullet/roman/…), the marker template (w:lvlText like "%1.%2"), the start value and a
// restart rule. The *displayed* number is stateful — it depends on how many same-level siblings
// precede it — so the running counters are computed by a pass over the document (model/render
// layer); this module supplies the definitions and the pure counters→string formatter.
import { findElement, findElements, getAttr } from "./xml";
import { parseMeasure } from "./units";
import { type RunProps, type ParagraphProps, parseRunProps, parseParagraphProps } from "./props";

export interface LevelDef {
  ilvl: number;
  start: number;
  /** w:numFmt: "decimal", "bullet", "lowerRoman", "upperLetter", "none", … */
  numFmt: string;
  /** w:lvlText: marker template, e.g. "%1.%2)" or a literal bullet glyph for bullets. */
  lvlText: string;
  isBullet: boolean;
  /** w:suff: separator after the marker. */
  suffix: "tab" | "space" | "nothing";
  /** 1-based level whose change restarts this level's counter (w:lvlRestart). */
  lvlRestart?: number;
  pPr: ParagraphProps;
  rPr: RunProps;
}

interface NumDef {
  numId: number;
  abstractNumId: number;
  /** Per-level overrides: ilvl → { start?, lvl? }. */
  overrides: Map<number, { start?: number; lvl?: LevelDef }>;
}

export interface Numbering {
  abstractNums: Map<number, Map<number, LevelDef>>; // abstractNumId → (ilvl → level)
  nums: Map<number, NumDef>; // numId → num
}

function parseLevel(lvlXml: string, ilvlAttr: string | undefined): LevelDef {
  const inner = lvlXml;
  const numFmt = getAttr(findElement(inner, "w:numFmt")?.openTag ?? "", "w:val") ?? "decimal";
  const suff = getAttr(findElement(inner, "w:suff")?.openTag ?? "", "w:val");
  const pPrEl = findElement(inner, "w:pPr");
  const pPr = pPrEl ? parseParagraphProps(inner.slice(pPrEl.innerStart, pPrEl.innerEnd)) : {};
  const rPrEl = findElement(inner, "w:rPr", { from: pPrEl ? pPrEl.outerEnd : 0 });
  const rPr = rPrEl ? parseRunProps(inner.slice(rPrEl.innerStart, rPrEl.innerEnd)) : {};
  const restart = getAttr(findElement(inner, "w:lvlRestart")?.openTag ?? "", "w:val");
  return {
    ilvl: parseMeasure(ilvlAttr, 0),
    start: parseMeasure(getAttr(findElement(inner, "w:start")?.openTag ?? "", "w:val"), 1),
    numFmt,
    lvlText: getAttr(findElement(inner, "w:lvlText")?.openTag ?? "", "w:val") ?? "",
    isBullet: numFmt === "bullet",
    suffix: suff === "space" ? "space" : suff === "nothing" ? "nothing" : "tab",
    lvlRestart: restart !== undefined ? parseMeasure(restart) : undefined,
    pPr,
    rPr,
  };
}

/** Parse numbering.xml into abstractNum + num tables. */
export function parseNumbering(numberingXml: string): Numbering {
  const abstractNums = new Map<number, Map<number, LevelDef>>();
  for (const an of findElements(numberingXml, "w:abstractNum")) {
    const id = parseMeasure(getAttr(an.openTag, "w:abstractNumId"), NaN);
    if (!Number.isFinite(id)) continue;
    const body = numberingXml.slice(an.innerStart, an.innerEnd);
    const levels = new Map<number, LevelDef>();
    for (const lvl of findElements(body, "w:lvl")) {
      const level = parseLevel(body.slice(lvl.innerStart, lvl.innerEnd), getAttr(lvl.openTag, "w:ilvl"));
      levels.set(level.ilvl, level);
    }
    abstractNums.set(id, levels);
  }

  const nums = new Map<number, NumDef>();
  for (const n of findElements(numberingXml, "w:num")) {
    const numId = parseMeasure(getAttr(n.openTag, "w:numId"), NaN);
    if (!Number.isFinite(numId)) continue;
    const body = numberingXml.slice(n.innerStart, n.innerEnd);
    const abstractNumId = parseMeasure(getAttr(findElement(body, "w:abstractNumId")?.openTag ?? "", "w:val"), NaN);
    const overrides = new Map<number, { start?: number; lvl?: LevelDef }>();
    for (const ov of findElements(body, "w:lvlOverride")) {
      const ilvl = parseMeasure(getAttr(ov.openTag, "w:ilvl"), 0);
      const ovBody = body.slice(ov.innerStart, ov.innerEnd);
      const startOverride = getAttr(findElement(ovBody, "w:startOverride")?.openTag ?? "", "w:val");
      const lvlEl = findElement(ovBody, "w:lvl");
      overrides.set(ilvl, {
        start: startOverride !== undefined ? parseMeasure(startOverride) : undefined,
        lvl: lvlEl ? parseLevel(ovBody.slice(lvlEl.innerStart, lvlEl.innerEnd), getAttr(lvlEl.openTag, "w:ilvl")) : undefined,
      });
    }
    nums.set(numId, { numId, abstractNumId, overrides });
  }

  return { abstractNums, nums };
}

/** Resolve the level definition for a (numId, ilvl), applying any w:lvlOverride. */
export function getLevel(numbering: Numbering, numId: number, ilvl: number): LevelDef | undefined {
  const num = numbering.nums.get(numId);
  if (!num) return undefined;
  const override = num.overrides.get(ilvl);
  if (override?.lvl) return override.lvl;
  const base = numbering.abstractNums.get(num.abstractNumId)?.get(ilvl);
  if (!base) return undefined;
  if (override?.start !== undefined) return { ...base, start: override.start };
  return base;
}

// ── integer → marker-glyph conversions ───────────────────────────────────────────────────────────

/** 1→I, 4→IV, 9→IX, 2026→MMXXVI. Non-positive input falls back to the decimal string. */
export function toRoman(n: number): string {
  if (n <= 0) return String(n);
  const table: [number, string][] = [
    [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"],
    [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
  ];
  let out = "";
  let rem = Math.floor(n);
  for (const [v, s] of table) while (rem >= v) { out += s; rem -= v; }
  return out;
}

/** Bijective base-26: 1→A, 26→Z, 27→AA. Non-positive input falls back to the decimal string. */
export function toLetter(n: number): string {
  if (n <= 0) return String(n);
  let out = "";
  let x = Math.floor(n);
  while (x > 0) {
    const rem = (x - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    x = Math.floor((x - 1) / 26);
  }
  return out;
}

function formatOne(value: number, numFmt: string): string {
  switch (numFmt) {
    case "decimal": return String(value);
    case "decimalZero": return value < 10 ? "0" + value : String(value);
    case "lowerLetter": return toLetter(value).toLowerCase();
    case "upperLetter": return toLetter(value);
    case "lowerRoman": return toRoman(value).toLowerCase();
    case "upperRoman": return toRoman(value);
    case "none": return "";
    default: return String(value);
  }
}

/**
 * Render the marker string for a list paragraph. `counters[k]` is the current 1-based count at
 * level k. For a bullet level the marker is the literal lvlText glyph; otherwise each "%N" in the
 * template is replaced by counters[N-1] formatted with level N-1's own numFmt.
 */
export function formatMarker(numbering: Numbering, numId: number, ilvl: number, counters: number[]): string {
  const level = getLevel(numbering, numId, ilvl);
  if (!level) return "";
  if (level.isBullet) return level.lvlText;
  return level.lvlText.replace(/%(\d+)/g, (_m, d: string) => {
    const lvlIndex = parseInt(d, 10) - 1;
    const lvl = getLevel(numbering, numId, lvlIndex);
    const value = counters[lvlIndex] ?? lvl?.start ?? 1;
    return formatOne(value, lvl?.numFmt ?? "decimal");
  });
}
