import { describe, it, expect } from "vitest";
import {
  findElement,
  findElements,
  childElements,
  replaceSpan,
  replaceInner,
  replaceNthElement,
  getAttr,
  parseAttrs,
  escapeXmlText,
  escapeXmlAttr,
  unescapeXml,
  localName,
} from "../src/xml";

describe("xml: OOXML element scanner", () => {
  it("locates an element's outer and inner spans", () => {
    const xml = `<w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body>`;
    const p = findElement(xml, "w:p")!;
    expect(p).toBeTruthy();
    expect(xml.slice(p.outerStart, p.outerEnd)).toBe(`<w:p><w:r><w:t>hi</w:t></w:r></w:p>`);
    expect(xml.slice(p.innerStart, p.innerEnd)).toBe(`<w:r><w:t>hi</w:t></w:r>`);
    expect(p.selfClosing).toBe(false);
  });

  it("depth-counts nested elements of the same name (table in a table)", () => {
    const xml = `<root><w:tbl>A<w:tbl>B</w:tbl>C</w:tbl><w:tbl>D</w:tbl></root>`;
    const first = findElement(xml, "w:tbl")!;
    expect(xml.slice(first.innerStart, first.innerEnd)).toBe(`A<w:tbl>B</w:tbl>C`);
    // The next sibling table (nth:1) is the OUTER second one, not the nested inner one.
    const second = findElement(xml, "w:tbl", { nth: 1 })!;
    expect(xml.slice(second.innerStart, second.innerEnd)).toBe(`D`);
  });

  it("handles self-closing elements", () => {
    const xml = `<w:p><w:pPr/><w:r/></w:p>`;
    const pPr = findElement(xml, "w:pPr")!;
    expect(pPr.selfClosing).toBe(true);
    expect(pPr.innerStart).toBe(pPr.innerEnd);
    expect(pPr.outerStart).toBe(xml.indexOf("<w:pPr/>"));
  });

  it("respects quoted attribute values that contain '>'", () => {
    const xml = `<w:p w:x="a>b"><w:t>x</w:t></w:p>`;
    const p = findElement(xml, "w:p")!;
    expect(p.openTag).toBe(`<w:p w:x="a>b">`);
    expect(xml.slice(p.innerStart, p.innerEnd)).toBe(`<w:t>x</w:t>`);
    expect(getAttr(p.openTag, "w:x")).toBe("a>b");
  });

  it("skips comments / CDATA / PIs so markup inside them isn't matched", () => {
    const xml = `<!-- <w:p>fake</w:p> --><?pi <w:p/>?><w:p>real</w:p>`;
    const p = findElement(xml, "w:p")!;
    expect(xml.slice(p.innerStart, p.innerEnd)).toBe("real");
  });

  it("findElements returns every sibling in a flat list", () => {
    const xml = `<w:styles><w:style w:styleId="a"/><w:style w:styleId="b"><x/></w:style></w:styles>`;
    const styles = findElements(xml, "w:style");
    expect(styles).toHaveLength(2);
    expect(getAttr(styles[0].openTag, "w:styleId")).toBe("a");
    expect(getAttr(styles[1].openTag, "w:styleId")).toBe("b");
  });

  it("childElements returns direct children of any name, not nested ones", () => {
    const xml = `<w:body><w:p>1</w:p><w:tbl><w:tr><w:p>nested</w:p></w:tr></w:tbl><w:p/></w:body>`;
    const body = findElement(xml, "w:body")!;
    const kids = childElements(xml, body.innerStart, body.innerEnd);
    expect(kids.map((k) => k.name)).toEqual(["w:p", "w:tbl", "w:p"]); // the in-table <w:p> is NOT a direct child
    expect(kids[2].selfClosing).toBe(true);
  });

  it("findElement honours the `to` bound (no leaking into siblings)", () => {
    const xml = `<a><w:p><w:pPr><w:b/></w:pPr></w:p><w:p><w:i/></w:p></a>`;
    const first = findElement(xml, "w:p")!;
    // Within the first paragraph only: it has a pPr; searching bounded must not find the 2nd p's content.
    expect(findElement(xml, "w:pPr", { from: first.innerStart, to: first.innerEnd })).toBeTruthy();
    expect(findElement(xml, "w:i", { from: first.innerStart, to: first.innerEnd })).toBeUndefined();
  });
});

describe("xml: surgical patching (preserve-and-patch)", () => {
  it("replaceSpan swaps one element and leaves every other byte intact", () => {
    const xml = `<w:body><w:p>one</w:p><w:p>two</w:p><w:p>three</w:p></w:body>`;
    const out = replaceNthElement(xml, "w:p", 1, `<w:p>TWO</w:p>`);
    expect(out).toBe(`<w:body><w:p>one</w:p><w:p>TWO</w:p><w:p>three</w:p></w:body>`);
  });

  it("replaceInner rewrites content only", () => {
    const xml = `<w:p w:rsidR="00A"><w:r><w:t>old</w:t></w:r></w:p>`;
    const span = findElement(xml, "w:p")!;
    const out = replaceInner(xml, span, `<w:r><w:t>new</w:t></w:r>`);
    expect(out).toBe(`<w:p w:rsidR="00A"><w:r><w:t>new</w:t></w:r></w:p>`);
  });

  it("replaceInner expands a self-closing element to hold content", () => {
    const xml = `<a><w:pPr/></a>`;
    const span = findElement(xml, "w:pPr")!;
    expect(replaceInner(xml, span, `<w:jc w:val="center"/>`)).toBe(
      `<a><w:pPr><w:jc w:val="center"/></w:pPr></a>`,
    );
  });

  it("a no-op patch of the nth element is identity", () => {
    const xml = `<w:body><w:p>one</w:p><w:p>two</w:p></w:body>`;
    const span = findElement(xml, "w:p", { nth: 0 })!;
    expect(replaceSpan(xml, span, xml.slice(span.outerStart, span.outerEnd))).toBe(xml);
  });
});

describe("xml: attributes + escaping", () => {
  it("getAttr matches whole attribute names, not suffixes", () => {
    const open = `<w:tcW w:w="1234" w:type="dxa">`;
    expect(getAttr(open, "w:w")).toBe("1234");
    expect(getAttr(open, "w:type")).toBe("dxa");
    expect(getAttr(open, "w")).toBeUndefined(); // not a suffix of "w:w"
  });

  it("parseAttrs reads all attributes (name=value pairs only), single or double quoted", () => {
    expect(parseAttrs(`<w:p w:a='1' w:b="2"/>`)).toEqual({ "w:a": "1", "w:b": "2" });
  });

  it("escapes and unescapes round-trip", () => {
    expect(escapeXmlText(`a<b>c&d`)).toBe("a&lt;b&gt;c&amp;d");
    expect(escapeXmlAttr(`x"y'z`)).toBe("x&quot;y&apos;z");
    expect(unescapeXml("a&lt;b&gt;c&amp;d&#65;&#x42;")).toBe("a<b>c&dAB");
  });

  it("localName strips the namespace prefix", () => {
    expect(localName("w:p")).toBe("p");
    expect(localName("p")).toBe("p");
  });
});

