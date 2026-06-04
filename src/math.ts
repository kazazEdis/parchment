// Office Math (OMML, m: namespace) → MathML, which browsers render natively. Covers the common
// constructs (runs, fractions, sub/superscripts, radicals, delimiters, n-ary operators); anything
// unmodelled falls through to its child content so text is never lost. The original OMML is kept
// verbatim in the model (preserve-and-patch), so this is render-only.
import { childElements, findElement, findElements, getAttr, unescapeXml, localName } from "./xml";

const MATHML_NS = "http://www.w3.org/1998/Math/MathML";

const runText = (xml: string, from: number, to: number): string =>
  findElements(xml.slice(from, to), "m:t").map((t) => unescapeXml(xml.slice(from, to).slice(t.innerStart, t.innerEnd))).join("");

/** Wrap a run's text in mn (number), mo (operator), or mi (identifier). */
function leaf(text: string): string {
  if (!text) return "";
  if (/^\d+(\.\d+)?$/.test(text)) return `<mn>${text}</mn>`;
  if (/^[+\-*/=<>±×÷·,()[\]{}|]+$/.test(text)) return `<mo>${text}</mo>`;
  return `<mi>${text}</mi>`;
}

const inner = (xml: string, name: string, from: number, to: number): { s: number; e: number } | null => {
  const el = findElement(xml, name, { from, to });
  return el ? { s: el.innerStart, e: el.innerEnd } : null;
};

function convertChildren(xml: string, from: number, to: number): string {
  return childElements(xml, from, to).map((el) => convertEl(xml, el)).join("");
}

function part(xml: string, parentFrom: number, parentTo: number, name: string): string {
  const r = inner(xml, name, parentFrom, parentTo);
  return r ? `<mrow>${convertChildren(xml, r.s, r.e)}</mrow>` : "<mrow></mrow>";
}

function convertEl(xml: string, el: { name: string; innerStart: number; innerEnd: number }): string {
  const ln = localName(el.name);
  const { innerStart: s, innerEnd: e } = el;
  switch (ln) {
    case "r": return leaf(runText(xml, s, e).trim());
    case "f": return `<mfrac>${part(xml, s, e, "m:num")}${part(xml, s, e, "m:den")}</mfrac>`;
    case "sSup": return `<msup>${part(xml, s, e, "m:e")}${part(xml, s, e, "m:sup")}</msup>`;
    case "sSub": return `<msub>${part(xml, s, e, "m:e")}${part(xml, s, e, "m:sub")}</msub>`;
    case "sSubSup": return `<msubsup>${part(xml, s, e, "m:e")}${part(xml, s, e, "m:sub")}${part(xml, s, e, "m:sup")}</msubsup>`;
    case "rad": {
      const deg = inner(xml, "m:deg", s, e);
      const base = part(xml, s, e, "m:e");
      return deg && deg.e > deg.s ? `<mroot>${base}<mrow>${convertChildren(xml, deg.s, deg.e)}</mrow></mroot>` : `<msqrt>${base}</msqrt>`;
    }
    case "d": {
      const beg = "(";
      const end = ")";
      return `<mrow><mo>${beg}</mo>${part(xml, s, e, "m:e")}<mo>${end}</mo></mrow>`;
    }
    case "nary": {
      const op = (() => { const pr = findElement(xml, "m:chr", { from: s, to: e }); return pr ? getAttr(pr.openTag, "m:val") ?? "∑" : "∑"; })();
      return `<munderover><mo>${op}</mo>${part(xml, s, e, "m:sub")}${part(xml, s, e, "m:sup")}</munderover>${part(xml, s, e, "m:e")}`;
    }
    case "e": case "num": case "den": case "sup": case "sub": case "deg": case "oMath":
      return convertChildren(xml, s, e);
    default:
      return convertChildren(xml, s, e);
  }
}

/** Convert an OMML fragment (m:oMath or its inner) to a MathML <math> string. */
export function ommlToMathML(ommlXml: string): string {
  const root = findElement(ommlXml, "m:oMath");
  const body = root ? convertChildren(ommlXml, root.innerStart, root.innerEnd) : convertChildren(ommlXml, 0, ommlXml.length);
  return `<math xmlns="${MATHML_NS}">${body}</math>`;
}
