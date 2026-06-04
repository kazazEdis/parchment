// String-level OOXML utilities (ARCHITECTURE.md §3.1, §3.9).
//
// Why string-based and not a DOM: the fidelity strategy is "preserve-and-patch" — we keep each
// part's original bytes and rewrite only the elements the user edited, splicing the new subtree
// into the original XML so everything else stays byte-identical. A DOMParser/XMLSerializer round
// trip re-orders attributes and normalises whitespace, which defeats that. These helpers also run
// identically in the browser and in Node tests (no DOMParser dependency).
//
// The scanner is namespace-naive: it matches the *qualified* name as written ("w:p", "w:tbl"), so
// callers pass the prefix exactly as it appears in the part.

/** A located element span within a source string. All indices are absolute offsets into `xml`. */
export interface ElementSpan {
  /** Qualified name as written, e.g. "w:p". */
  name: string;
  /** Offset of the leading "<". */
  outerStart: number;
  /** Offset just past the closing ">" (of `</name>`, or of "/>" when self-closing). */
  outerEnd: number;
  /** Offset just past the open tag's ">". For self-closing elements, equals `outerEnd`. */
  innerStart: number;
  /** Offset of the closing "</name>"'s "<". For self-closing elements, equals `innerStart`. */
  innerEnd: number;
  /** The full open tag text, e.g. `<w:p w:rsidR="00A">`. For self-closing, includes the "/>". */
  openTag: string;
  /** True for `<w:p/>`-style empty elements. */
  selfClosing: boolean;
}

interface Token {
  kind: "open" | "close" | "selfclose" | "comment" | "cdata" | "pi" | "decl";
  start: number; // offset of "<"
  end: number; // just past the token's ">"
  name: string; // qualified element name for open/close/selfclose; "" otherwise
}

/** Read the next markup token at or after `from`. Returns null at end of input. */
function nextToken(xml: string, from: number): Token | null {
  const lt = xml.indexOf("<", from);
  if (lt < 0) return null;

  if (xml.startsWith("<!--", lt)) {
    const end = xml.indexOf("-->", lt + 4);
    return { kind: "comment", start: lt, end: end < 0 ? xml.length : end + 3, name: "" };
  }
  if (xml.startsWith("<![CDATA[", lt)) {
    const end = xml.indexOf("]]>", lt + 9);
    return { kind: "cdata", start: lt, end: end < 0 ? xml.length : end + 3, name: "" };
  }
  if (xml.startsWith("<?", lt)) {
    const end = xml.indexOf("?>", lt + 2);
    return { kind: "pi", start: lt, end: end < 0 ? xml.length : end + 2, name: "" };
  }
  if (xml.startsWith("<!", lt)) {
    const end = xml.indexOf(">", lt + 2);
    return { kind: "decl", start: lt, end: end < 0 ? xml.length : end + 1, name: "" };
  }

  // A real element tag. Walk to the matching ">", respecting quoted attribute values
  // (attribute values may legally contain a bare ">").
  const isClose = xml[lt + 1] === "/";
  let i = lt + (isClose ? 2 : 1);
  // name = up to whitespace, "/", or ">"
  const nameStart = i;
  while (i < xml.length && !/[\s/>]/.test(xml[i])) i++;
  const name = xml.slice(nameStart, i);

  let quote: string | null = null;
  while (i < xml.length) {
    const c = xml[i];
    if (quote) {
      if (c === quote) quote = null;
    } else if (c === '"' || c === "'") {
      quote = c;
    } else if (c === ">") {
      break;
    }
    i++;
  }
  const gt = i; // index of ">"
  const end = gt + 1;
  if (isClose) return { kind: "close", start: lt, end, name };
  const selfClosing = xml[gt - 1] === "/";
  return { kind: selfClosing ? "selfclose" : "open", start: lt, end, name };
}

/**
 * Find the `nth` (0-based) element named `name`, scanning from `from`. Correctly handles nested
 * elements of the same name (e.g. a table inside a table cell), self-closing tags, and skips
 * comment/CDATA/PI regions. Returns undefined if not found.
 */
export function findElement(
  xml: string,
  name: string,
  opts: { from?: number; to?: number; nth?: number } = {},
): ElementSpan | undefined {
  const target = opts.nth ?? 0;
  const limit = opts.to ?? xml.length;
  let seen = 0;
  let cursor = opts.from ?? 0;

  while (cursor < limit) {
    const tok = nextToken(xml, cursor);
    if (!tok || tok.start >= limit) return undefined;
    cursor = tok.end;

    if ((tok.kind === "open" || tok.kind === "selfclose") && tok.name === name) {
      if (seen < target) {
        // Skip this whole element (including nested same-name children) and keep counting.
        if (tok.kind === "selfclose") {
          seen++;
          continue;
        }
        const span = closeElement(xml, name, tok);
        seen++;
        cursor = span.outerEnd;
        continue;
      }
      // This is the one.
      if (tok.kind === "selfclose") {
        return {
          name,
          outerStart: tok.start,
          outerEnd: tok.end,
          innerStart: tok.end,
          innerEnd: tok.end,
          openTag: xml.slice(tok.start, tok.end),
          selfClosing: true,
        };
      }
      return closeElement(xml, name, tok);
    }
  }
  return undefined;
}

/** Given an open token, walk to its matching close tag (depth-aware) and build the span. */
function closeElement(xml: string, name: string, open: Token): ElementSpan {
  let depth = 1;
  let cursor = open.end;
  while (cursor < xml.length) {
    const tok = nextToken(xml, cursor);
    if (!tok) break;
    cursor = tok.end;
    if (tok.name === name) {
      if (tok.kind === "open") depth++;
      else if (tok.kind === "close") {
        depth--;
        if (depth === 0) {
          return {
            name,
            outerStart: open.start,
            outerEnd: tok.end,
            innerStart: open.end,
            innerEnd: tok.start,
            openTag: xml.slice(open.start, open.end),
            selfClosing: false,
          };
        }
      }
      // selfclose of same name does not change depth
    }
  }
  // Unbalanced input: treat the rest of the document as the body (defensive, shouldn't happen).
  return {
    name,
    outerStart: open.start,
    outerEnd: xml.length,
    innerStart: open.end,
    innerEnd: xml.length,
    openTag: xml.slice(open.start, open.end),
    selfClosing: false,
  };
}

/**
 * All non-overlapping elements named `name`, scanned left-to-right. After each match the scan
 * resumes past that element's end, so this returns siblings at the scanned level but not matches
 * nested *inside* a returned element. Ideal for flat lists (e.g. every `w:style` in styles.xml).
 */
export function findElements(xml: string, name: string, from = 0): ElementSpan[] {
  const out: ElementSpan[] = [];
  let cursor = from;
  while (cursor < xml.length) {
    const span = findElement(xml, name, { from: cursor });
    if (!span) break;
    out.push(span);
    cursor = span.outerEnd;
  }
  return out;
}

/**
 * All **direct child** elements within `[from, to)`, in document order, regardless of name. The
 * tree-walking primitive: after each child its whole subtree is skipped, so nested elements are not
 * returned (unlike {@link findElements}, which filters by name but descends into non-matching
 * elements). Text/comment/PI nodes between children are ignored. Pass a `w:body`/`w:tr` inner span
 * to enumerate its blocks/cells.
 */
export function childElements(xml: string, from = 0, to: number = xml.length): ElementSpan[] {
  const out: ElementSpan[] = [];
  let cursor = from;
  while (cursor < to) {
    const tok = nextToken(xml, cursor);
    if (!tok || tok.start >= to) break;
    cursor = tok.end;
    if (tok.kind === "open") {
      const span = closeElement(xml, tok.name, tok);
      out.push(span);
      cursor = span.outerEnd;
    } else if (tok.kind === "selfclose") {
      out.push({
        name: tok.name,
        outerStart: tok.start,
        outerEnd: tok.end,
        innerStart: tok.end,
        innerEnd: tok.end,
        openTag: xml.slice(tok.start, tok.end),
        selfClosing: true,
      });
    }
    // close/comment/cdata/pi/decl tokens are skipped.
  }
  return out;
}

/** Replace the element at `span` with `replacement`, leaving every other byte untouched. */
export function replaceSpan(xml: string, span: ElementSpan, replacement: string): string {
  return xml.slice(0, span.outerStart) + replacement + xml.slice(span.outerEnd);
}

/** Replace the inner content of `span` (between its open and close tags) with `inner`. */
export function replaceInner(xml: string, span: ElementSpan, inner: string): string {
  if (span.selfClosing) {
    // Expand `<w:p/>` into `<w:p>…</w:p>` so it can hold content.
    const openNoSlash = span.openTag.replace(/\/>$/, ">");
    return xml.slice(0, span.outerStart) + openNoSlash + inner + `</${span.name}>` + xml.slice(span.outerEnd);
  }
  return xml.slice(0, span.innerStart) + inner + xml.slice(span.innerEnd);
}

/** Convenience: replace the nth (0-based) `name` element's outer XML. No-op if not found. */
export function replaceNthElement(xml: string, name: string, nth: number, replacement: string): string {
  const span = findElement(xml, name, { nth });
  return span ? replaceSpan(xml, span, replacement) : xml;
}

// ── attributes ────────────────────────────────────────────────────────────────────────────────

/** Read one attribute value (unescaped) from an open-tag string. Undefined if absent. */
export function getAttr(openTag: string, attr: string): string | undefined {
  // Match the attribute name as a whole token to avoid matching a suffix of a longer name.
  const re = new RegExp(`(?:^|[\\s])${escapeRegExp(attr)}\\s*=\\s*("([^"]*)"|'([^']*)')`);
  const m = re.exec(openTag);
  if (!m) return undefined;
  return unescapeXml(m[2] ?? m[3] ?? "");
}

/** Parse all attributes of an open-tag string into a map (values unescaped). */
export function parseAttrs(openTag: string): Record<string, string> {
  const out: Record<string, string> = {};
  const re = /([\w:.-]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(openTag)) !== null) {
    out[m[1]] = unescapeXml(m[3] ?? m[4] ?? "");
  }
  return out;
}

// ── escaping ──────────────────────────────────────────────────────────────────────────────────

/** Escape text content for XML (`&`, `<`, `>`). */
export function escapeXmlText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Escape an attribute value for XML (text escapes plus quotes). */
export function escapeXmlAttr(s: string): string {
  return escapeXmlText(s).replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

/** Decode the five predefined XML entities plus numeric (&#NN; / &#xNN;) references. */
export function unescapeXml(s: string): string {
  return s.replace(/&(#x?[0-9a-fA-F]+|amp|lt|gt|quot|apos);/g, (whole, body: string) => {
    switch (body) {
      case "amp": return "&";
      case "lt": return "<";
      case "gt": return ">";
      case "quot": return '"';
      case "apos": return "'";
      default: {
        const code = body[1] === "x" || body[1] === "X"
          ? parseInt(body.slice(2), 16)
          : parseInt(body.slice(1), 10);
        return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
      }
    }
  });
}

/** The local name of a qualified name ("w:p" → "p", "p" → "p"). */
export const localName = (qname: string): string => {
  const i = qname.indexOf(":");
  return i < 0 ? qname : qname.slice(i + 1);
};

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
