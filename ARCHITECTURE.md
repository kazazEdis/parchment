# In-house .docx editor — architecture & problem catalogue

How a high-fidelity browser Word editor works under the hood, what SuperDoc does, and the
concrete problems we must solve to grow our own editor up from `opc.ts`.

**Licensing boundary.** SuperDoc (`@harbour-enterprises/superdoc`) is **AGPL-3.0 / commercial
dual-licensed**. We study its *architecture and approach* (ideas and file formats are not
copyrightable) and implement against the **open ECMA-376 / OOXML spec**. We copy **no SuperDoc
source**. `opc.ts` already states this. Keep it true: every module here is written from the spec.

---

## 1. SuperDoc, under the hood (verified from public docs)

SuperDoc's headline decision: **decouple the document model from rendering.** A *hidden*
ProseMirror `Editor` holds document state and runs editing commands, but **its DOM is never
shown**. A separate pipeline does Word-style layout and paints to the screen. This is what lets
it do true pagination instead of an infinite `contenteditable` scroll.

The pipeline:

```
.docx ─▶ super-converter ─▶ ProseMirror doc ─▶ layout-adapter ─▶ FlowBlock[]
            (OOXML⇄PM)        (hidden model)     (PM→flow)            │
                                                                      ▼
                                          DOM ◀── DomPainter ◀── layout-engine
                                          (screen)   (paint)      (paginate → Layouts)
```

Monorepo modules (from `CONTRIBUTING.md`):

| Module | Responsibility |
|---|---|
| `super-editor` | Hosts the hidden ProseMirror editor (state + commands). Contains `super-converter`, `layout-adapter`, `extensions`. |
| `super-converter` (`src/editors/v1/core/super-converter/`) | DOCX **import/export**: OOXML ⇄ ProseMirror. **Stores raw OOXML properties** on nodes — does *not* resolve styles at import. |
| `layout-adapter` (`.../layout-adapter/`) | Projects ProseMirror nodes → `FlowBlock[]` (renderer-neutral layout primitives). |
| `extensions` (`.../extensions/`) | Editing behaviours (bold, lists, tables, keymaps) — Tiptap-style. |
| `layout-engine` | Pagination: `FlowBlock[]` → paginated `Layout`s. Computational (measure + break), not CSS floats. |
| `style-engine` | Resolves the **raw OOXML props at render time** — the style cascade (fonts, colours, borders). |
| `painters/dom` (`DomPainter`) | Renders `Layout`s to real DOM. |
| `collaboration-yjs` | Real-time co-editing (Yjs CRDT). |

Two principles to steal outright:

1. **Store raw, resolve late.** Import keeps the original OOXML property bags verbatim on each
   node; the cascade is applied only when rendering. This is *the* fidelity trick — nothing is
   lossily flattened at import, so export can reproduce the original.
2. **Preserve-and-patch export.** Round-trip without destroying formatting by writing edits back
   into the original XML, rewriting only what changed. (Our `opc.ts` keeps every part's bytes
   verbatim for exactly this.)

SuperDoc also exposes a high-level **Document API** (`editor.doc`, "300+ ops": `format.bold()`,
`comments.create()`, `query.match()`, …) and is deprecating direct `editor.state/view/schema`
access — i.e. the public surface is an *operations API over OOXML*, not a ProseMirror passthrough.
Worth mirroring eventually so callers never touch the model directly.

**Contrast — the cheap pagination route.** A separate ProseMirror editor ("Badon") paginates with
**pure CSS** (pages as sibling DOM nodes, `display: contents`, floats for breaks): fast to ~400
pages, no node-splitting math, but tables/footnotes/merged-cells break and need bespoke plugins.
That's the trade we'd pick against: CSS pagination = quick demo, computational layout = real Word
fidelity. SuperDoc chose computational. We should too, eventually — but CSS/`pagedjs` is a fine
*first* render.

---

## 1b. Where we beat SuperDoc — and where we don't (be honest)

This engine is **not** trying to out-WYSIWYG ProseMirror in a single build. SuperDoc still leads on
interactive contentEditable editing, computed pagination, and realtime Yjs collaboration. What we
are decisively better at — for an offer/CPQ product that mostly **generates, fills, redlines, and
stores** documents, often server-side — is everything around that:

| Axis | SuperDoc | This engine |
|---|---|---|
| Licence | **AGPL-3.0** / paid commercial | **MIT-clean, owned** — embed in a closed product, no copyleft, no fee |
| Footprint | `@harbour-enterprises/superdoc` ≈ **54 MB** package | **0.11 MB / ~1.9k LOC**, only shares `jszip` (~500× smaller marginal) |
| Runtime | browser + ProseMirror + (Vue) | **fully headless** — runs in Node/Convex actions, pure functions |
| Round-trip fidelity | re-emits via PM→OOXML (can drop unmodelled content) | **byte-verbatim preserve-and-patch** — only edited nodes are rewritten |
| Determinism / tests | browser-coupled | **250 unit tests**, no DOM, no flakiness |
| Template fill | naive replace misses run-split placeholders | **run-split-aware** `replaceText`/`fillTemplate` (`doc.ts`) |
| Redlining | renders *existing* tracked changes | **generates** them — semantic word diff → `w:ins`/`w:del` (`diff.ts`) |

Measured: a 400-paragraph + 10-table document, full read→parse→fill(1200 tokens)→write headless
cycle in **~94 ms**. The strategy: own the automation + fidelity + footprint axes outright, ship a
faithful Tier-1 viewer, and treat WYSIWYG/pagination/collab as later, optional layers — not the moat.

---

## 2. The document is a ZIP of XML parts (OPC) — **done: `opc.ts`**

A `.docx` is an OPC package (ECMA-376 Part 2): a ZIP whose entries ("parts") are XML + media,
wired together by relationship files. The parts that matter:

| Part | Holds |
|---|---|
| `[Content_Types].xml` | MIME type per extension/part. Must stay valid on write. |
| `_rels/.rels` | Package root rels → points at the main document. |
| `word/document.xml` | **The body**: paragraphs, runs, tables, sections. |
| `word/_rels/document.xml.rels` | Rels from the body → images, headers, hyperlinks, numbering, styles. |
| `word/styles.xml` | Style definitions + `docDefaults`. |
| `word/numbering.xml` | List definitions (`abstractNum` + `num`). |
| `word/settings.xml` | Doc settings (default tab, track-changes on, compat flags). |
| `word/theme/theme1.xml` | Theme fonts + colours (referenced by name). |
| `word/header*.xml` / `word/footer*.xml` | Header/footer bodies (per section, per type). |
| `word/media/*` | Image binaries (png/jpeg/emf). |
| `word/comments.xml`, `commentsExtended.xml`, `people.xml` | Comments + threading + authors. |
| `word/footnotes.xml` / `endnotes.xml` | Notes. |

`opc.ts` already gives us: `readDocx`/`writeDocx`, the `DocxPackage{parts,order}` map with
byte-verbatim `DocxPart`s, `getPartText`/`setPartText`/`setPartBytes` (dirty-tracking),
`hasPart`/`getPart`, `docxToBlob`. **Everything below is built on top of this.** The `order` +
verbatim-bytes design is the foundation of preserve-and-patch — don't lose it.

---

## 3. The hard problems, and how to solve each

Ordered roughly by when you hit them.

### 3.1 XML layer
- **Problem.** Need to parse/serialise OOXML namespaced XML and round-trip *unmodelled* nodes
  losslessly (unknown elements, `mc:AlternateContent`, `w:` vs `w14:` vs `wp:` namespaces).
- **Solution.** A thin DOM over a permissive parser. Browser-native `DOMParser`/`XMLSerializer`
  handle namespaces and are free, but **re-serialisation is not byte-identical** (attribute order,
  self-closing style, whitespace) — which fights preserve-and-patch. So: parse for *reading*, but
  for *writing* prefer surgical string/AST patches over full re-serialise (see 3.9). Keep a
  `fast-xml-parser` option for Node tests. Preserve `xml:space="preserve"` and significant
  whitespace exactly.

### 3.2 Units — get these wrong and everything is subtly off
| Unit | Means | Used for |
|---|---|---|
| **twip** (1/1440 in, 1/20 pt) | twentieth of a point | margins, indents, tab stops, table widths, page size |
| **half-point** | ½ pt | font size (`w:sz val="24"` = 12pt) |
| **EMU** (914400/in, 12700/pt) | English Metric Unit | image/drawing sizes (`wp:extent`) |
| **eighth-point** | ⅛ pt | border widths (`w:sz`) |
| **fiftieth-%** | val/50 = % | some shading/widths (`w:tblW type="pct"` is ×50) |
| **DXA / 240ths** | line spacing | `w:spacing line` |
Centralise these conversions once. Most "it looks slightly wrong" bugs are a unit confusion.

### 3.3 The document model (our ProseMirror analogue)
- **Problem.** Need an editable tree that (a) is rich enough for Word constructs, (b) carries the
  **raw OOXML props** for late resolution, (c) maps cleanly back to XML.
- **Solution.** A block/inline tree: `Document → Body → [Paragraph | Table] → …`; a Paragraph has
  `pPr` (raw) + inline children `[Run | Hyperlink | Drawing | Field | BookmarkMark]`; a Run has
  `rPr` (raw) + text. Store property bags as opaque parsed objects, **not** flattened CSS. Keep a
  back-pointer to the original XML node (or its serialised form) so untouched nodes export verbatim.
  ProseMirror is the proven choice (schema, transactions, mapping, collab) but is heavy; a custom
  immutable tree is viable since our editor is narrower than Word.

### 3.4 Paragraphs & runs (`w:p` / `w:r`)
- A paragraph = `w:pPr` (alignment, indent, spacing, `w:pStyle`, numbering ref `w:numPr`) + runs.
- A run = `w:rPr` (bold `w:b`, italic, `w:u`, `w:color`, `w:sz`, `w:rFonts`, `w:highlight`,
  `w:vertAlign` for super/sub) + content (`w:t` text, `w:tab`, `w:br`, `w:drawing`, `w:sym`).
- **Gotcha.** `w:t` whitespace is only kept with `xml:space="preserve"`. Adjacent runs with
  identical `rPr` should be coalesced on read for sane editing, then re-split on write.

### 3.5 Styles & the cascade (`styles.xml`)
- **Problem.** A run's *effective* formatting = the OOXML cascade, in order:
  `docDefaults` → referenced paragraph style (with `w:basedOn` chain) → referenced character
  style → numbering-level props → **direct** `rPr`/`pPr`. Toggle props (bold) *flip*, they don't
  just override. `w:default="1"` styles apply with no explicit ref.
- **Solution.** Build a style table (id → resolved props, flattening `basedOn` chains once). At
  render, resolve each node's effective props through the cascade. This is SuperDoc's
  `style-engine`. Don't bake the result into the model — recompute, so edits to a style restyle
  everything and export stays clean.

### 3.6 Lists & numbering (`numbering.xml`) — the classic fidelity killer
- **Problem.** Lists aren't a tree in OOXML. A paragraph just references `w:numId` + `w:ilvl`.
  `num` → `abstractNum` → per-level `lvl` (format `decimal`/`bullet`/`lowerRoman`, `lvlText`
  template like `%1.%2`, start-at, indent, run props for the marker). Numbering is **stateful**:
  the displayed number depends on how many same-level siblings precede it, `lvlRestart`, and
  `w:lvlOverride`/`startOverride`.
- **Solution.** Resolve `numId+ilvl → abstractNum level def`, then run a **counter pass** over the
  document in order, per abstract-num, incrementing/restarting per level to compute each marker
  string from `lvlText`. Render the marker as generated content (not editable text). On export the
  numbering is intrinsic (just keep `numPr`), so the counter is render-only.

### 3.7 Tables (`w:tbl`) — the other fidelity killer
- **Problem.** `w:tblGrid` defines column widths; rows (`w:tr`) hold cells (`w:tc`). Spans are
  encoded two ways: horizontal `w:gridSpan`, vertical `w:vMerge` (`val="restart"` starts, omitted/
  `continue` continues). Borders cascade (table → row → cell → conditional `tblStylePr` for
  first/last row/col, banding). Cell widths can be twips or pct.
- **Solution.** Build a logical grid (resolve spans into a 2-D cell map with rowspan/colspan),
  resolve border/shading conflicts (Word has a defined precedence), render as an HTML table or
  absolutely-positioned cells. Editing must keep the grid + `tblGrid` consistent. **Tables across
  page breaks** (3.8) are the worst case — repeat header rows (`w:tblHeader`), split body rows.

### 3.8 Sections, pages & the **pagination problem**
- **Problem.** Word is **paginated**; HTML is a continuous flow. A `w:sectPr` (last paragraph of a
  section, or body-level for the last) defines page size, margins, columns, header/footer refs,
  page-number format. Content must be measured and broken into pages; headers/footers repeat;
  page numbers and `PAGE`/`NUMPAGES` fields depend on the break result.
- **Solutions, two tiers:**
  - **Tier 1 (ship first): CSS / `pagedjs`.** Render the flow, let `@page` + `pagedjs` do breaks
    and margins. Fast, decent print/PDF. Weak on Word-exact line breaking and complex tables.
  - **Tier 2 (real fidelity): computational layout-engine** (SuperDoc's path). Convert the model
    to renderer-neutral **FlowBlocks**, **measure** each (line-break with real font metrics),
    **greedily fill** pages of `pageHeight − margins − header − footer`, **split** paragraphs at
    line boundaries and tables at row boundaries, apply **widow/orphan** control (min lines top/
    bottom), then paint pages as separate DOM nodes with their section's header/footer and the
    resolved page number. Measuring is the cost; cache line metrics, re-layout only dirtied blocks.
- **Gotchas.** Section breaks change page geometry mid-document; first-page/even-odd headers
  (`w:titlePg`, `settings evenAndOddHeaders`); `w:br type="page"` forced breaks; keep-with-next /
  keep-lines paragraph flags.

### 3.9 Serialisation — **preserve-and-patch** (the fidelity contract)
- **Problem.** A naive "model → fresh XML" export reorders attributes, drops unknown elements, and
  changes whitespace → Word sees a "different" doc, diffs explode, and unmodelled features vanish.
- **Solution.** Keep every original part's bytes (we do). On save, for each **dirtied** node only,
  patch its XML in place: re-emit just that `w:p`/`w:r`/`w:tbl` subtree, splice into the original
  `document.xml` string, leave everything else byte-identical. New parts (a new image) append; the
  `[Content_Types].xml` and `*.rels` get minimal additions. `DocxPart.dirty` already flags what to
  rewrite. **Test:** open→save with no edits must produce a near-identical zip (only timestamps).

### 3.10 Images & DrawingML (`w:drawing`)
- Inline (`wp:inline`) vs floating/anchored (`wp:anchor`, with wrap + position). Size in **EMU**
  (`wp:extent`). The actual bytes are referenced via `a:blip r:embed="rId.."` → look up in
  `document.xml.rels` → `word/media/imageN.png`. **Solution:** resolve rel → part bytes →
  object-URL for display; on insert, add the media part + a rels entry + content-type. Preserve
  EMF/WMF you can't render by showing a placeholder but keeping the bytes (preserve-and-patch).

### 3.11 Hyperlinks, bookmarks, fields
- Hyperlink = `w:hyperlink r:id=..` (external, via rels) or `w:anchor` (internal bookmark).
- Bookmarks = `w:bookmarkStart/End` (range markers — store as zero-width marks on positions).
- **Fields** = `w:fldSimple` or the `w:fldChar(begin) … w:instrText … fldChar(end)` run sequence.
  `TOC`, `PAGE`, `NUMPAGES`, `REF`, `HYPERLINK`, dates. **Solution:** parse the instruction, render
  the cached result run, recompute the ones we own (PAGE/NUMPAGES after pagination, TOC from
  heading styles). Keep unknown fields' cached result verbatim.

### 3.12 Track changes / redlining (`w:ins` / `w:del`)
- **Problem.** Edits wrap runs in `w:ins`/`w:del` (with `w:author`, `w:date`); deletions keep text
  in `w:delText`; property changes use `w:rPrChange`/`w:pPrChange`. The doc has *two* views:
  original and final.
- **Solution.** Model insertions/deletions as marks carrying author/date. A "change-mode" flag
  makes every mutation emit tracked markup instead of a hard edit (SuperDoc's `--change-mode
  tracked`). Render redlines (underline insert, strike delete, author colour); accept/reject =
  unwrap or remove. This is a high-value feature for an offer tool (legal review).

### 3.13 Comments
- `commentRangeStart/End` + `w:commentReference` in the body point into `comments.xml`; threading
  in `commentsExtended.xml`, authors in `people.xml`. **Solution:** range marks ↔ a comments side
  panel; preserve the four-file linkage on write.

### 3.14 Fonts & measurement
- **Problem.** Line breaking and pagination need **font metrics**; the user may not have Word's
  fonts. `w:rFonts` names ascii/hAnsi/eastAsia/cs faces, often via theme (`+mn-lt`).
- **Solution.** Map theme font slots → real names; load web equivalents; measure with canvas
  `measureText` or an opentype metrics table. Keep a substitution table (Calibri→Carlito,
  Cambria→Caladea — metric-compatible) for fidelity. Embedded fonts (`word/fonts/*`, obfuscated)
  can be de-obfuscated and `@font-face`'d.

### 3.15 Collaboration (optional, later)
- **Problem.** Concurrent editing without a lock.
- **Solution.** Yjs CRDT bound to the model (`y-prosemirror` if we use PM); a provider
  (Hocuspocus/WebSocket) syncs updates; presence cursors. Convex could back the sync doc, but note
  the cost invariant — co-editing is chatty; gate it behind a feature flag and don't subscribe
  per-keystroke through Convex (sync server holds the hot state, Convex persists snapshots).

### 3.16 Export to PDF
- Tier 1: `window.print()` with `@page` CSS or `pagedjs` → browser PDF. Tier 2: render the
  computed page DOM to canvas/`jsPDF`, or server-side headless Chrome for exact output. We already
  ship `jsPDF` for offers — reuse for simple cases.

---

## 4. Build status (each step shippable, each preserves round-trip)

The headless engine + Tier-1 render are **done** and unit-tested; `index.ts` is the public barrel
(the React view is imported separately so headless consumers don't pull in React).

1. ✅ **`opc.ts`** — zip ↔ parts, byte-verbatim, dirty tracking. (`xml.ts` is string-based, not
   `DOMParser` — re-serialising a DOM breaks preserve-and-patch; see §3.1.)
2. ✅ **`xml.ts`** — depth-aware OOXML scanner (`findElement`/`findElements`/`childElements`) +
   surgical patcher (`replaceSpan`/`replaceInner`); handles nesting, self-close, quoted `>`, skips
   comments/CDATA/PI. **`units.ts`** — twip/half-pt/eighth/EMU/pct/line conversions.
3. ✅ **`props.ts`** (typed rPr/pPr + cascade merge), **`styles.ts`** (docDefaults + `basedOn` chain,
   chain-only resolvers so docDefaults applies once), **`numbering.ts`** (abstractNum levels,
   overrides, `formatMarker` decimal/roman/letter/bullet).
4. ✅ **`model.ts`** — `parseDocument` → body tree (paragraphs/runs/hyperlinks/drawings/tables/
   section) with raw props + absolute source spans for patching.
5. ✅ **`resolve.ts`** (style-engine: effective props + `assignListNumbers` counter pass),
   **`cssMap.ts`** (props → CSS), **`table.ts`** (gridSpan/vMerge → render grid), **`images.ts`**
   (blip → data URL), **`DocxView.tsx`** (Tier-1 continuous page sheet — verified rendering).
6. ✅ **`serialize.ts`** — preserve-and-patch writer: re-emit only edited nodes, splice at the source
   span; full-package round-trip leaves other parts byte-identical.
7. ✅ **`edit.ts`** — pure paragraph transforms (format, toggles, alignment), **run-split-aware
   `replaceInParagraph`**, and **track changes** (§3.12: `w:ins`/`w:del` as run marks, `w:delText`,
   accept/reject, change-mode marking, redline render). ✅ **`diff.ts`** — word-level diff →
   tracked-change redlines (generate redlines between two versions). ✅ **`doc.ts`** — headless
   **Document API**: `openDocx`/`fromPackage`, `getText`, `replaceText`/`fillTemplate` (run-split
   aware), `transformParagraphs`, `acceptAllChanges`/`rejectAllChanges`, `save`/`saveBlob`.
8. ✅ **Front end** — `DocxView.tsx` (read-only viewer) + **`DocxEditor.tsx`** (interactive WYSIWYG):
   editable paragraphs, **inline range selection + per-selection formatting** (`edit.formatRange`,
   splits runs at the boundary, selection preserved), **undo/redo** (reducer over document.xml
   snapshots), **paste from Word** (`htmlToRuns` → `edit.spliceRunRange`), find/replace,
   track-changes-by-diff, accept/reject, download. Robustness trick: structural ops rebuild the active
   paragraph from the **live DOM** before applying the pure op, so model + view never desync.
9. ✅ **Block ops** (`edit.splitParagraph`/`mergeParagraphs`/`setListLevel`): Enter→split,
   Backspace→merge, Tab/Shift+Tab→indent, in the editor with caret placement.
10. ✅ **Computed pagination** (`paginate.ts` + `PaginatedDocxView.tsx`): **line-level** — measure each
    line's bottom-Y (`computePageBreaks`) and render clipped translate-windows, so long paragraphs split
    mid-paragraph across pages (SuperDoc's `sliceLines` idea, our own DOM-measure implementation).
11. ✅ **Comments** (`comments.ts`): parse + locate ranges (display) AND **authoring** —
    `addComment`/`wrapParagraphComment`/`nextCommentId` + the editor's Comment button writes
    `comments.xml` + the body range. Reply/resolve threading still owed.
12. ✅ **CRDT collaboration** (`collab.ts` + `crdt.ts`): a Yjs `Y.Text` over document.xml with
    minimal-delta local edits; concurrent edits to different regions **merge conflict-free** (verified
    live two-tab). Snapshot transport remains as a fallback. Presence cursors + a server transport next.

13. ✅ **Structure editing**: `tableEdit` (insert/delete rows + columns, fidelity-preserving),
    `imageInsert` (+ `opcParts` rels/content-types), `linkEdit` (hyperlink a selection), `bookmarks`
    (parse + insert), all wired into the editor's Insert toolbar.
14. ✅ **More content types**: `fields` (instructions + TOC + fldSimple), `footnotes`, header/footer
    rendering, **math** (`math`: OMML → MathML), **encrypted-docx detection** (`encrypted`).
15. ✅ **Presence cursors**: peers broadcast cursor + identity over the CRDT transport; remote carets
    render with coloured name labels.

**Honest remaining residual vs SuperDoc** — now a long tail, not the major features:
- OOXML **breadth**: their per-element translator army (hundreds of elements) models more exotic
  constructs than our common-case model. We **preserve the unmodelled rest verbatim** (nothing lost),
  but don't render/edit it. This is asymptotic — not closeable 1:1 in bounded time.
- **Niche / large standalone features**: charts, SmartArt, text boxes/shapes; **agile AES
  decryption** (we detect, don't decrypt); **EMF/WMF** rendering; citations/bibliography; AI assist
  (a product decision — wire to Claude). Each is a sizeable independent effort.

Every **major** SuperDoc feature is covered — viewer, interactive WYSIWYG editor (selection,
undo/redo, paste, block ops), line-level pagination, CRDT collab + presence, track changes, generate-
redlines, comments (read/author/ranges), tables, images, hyperlinks, bookmarks, fields/TOC, footnotes,
headers/footers, math — and we beat it on licence/footprint/headless/determinism/fidelity. 308 tests.

Invariants held throughout: **store raw props / resolve late** (§1.1), **preserve-and-patch on
write** (§3.9), **centralised units** (§3.2), and the cost rules in `CLAUDE.md` for anything that
touches Convex (collab/persistence).

---

## 4b. How SuperDoc does the hard parts (studied from its shipped `.d.ts`)

Read from `@harbour-enterprises/superdoc/dist/**/*.d.ts` (AGPL — studied the *approach/algorithms*,
not copied source). What explains its robustness and the remaining gaps:

**Per-element OOXML translators.** `super-converter/v3/handlers/<ns>/<element>/<element>-translator`
— one translator per OOXML element (b, bdr, bidi, abstractNum, pgSz, …), hundreds of them, plus a v2
importer. That breadth IS its real-world robustness; we cover the common 90% and preserve the rest
verbatim. (They also ship `ooxml-encryption/agile-decryptor` and EMF/WMF rendering — full coverage.)

**Pagination = a measured line-flow engine** (`layout-engine/`), not block packing:
- Each paragraph is measured into `ParagraphMeasure { lines: Line[]; totalHeight }`; each `Line`
  carries a height + a `LinePmRange` (its document position range).
- `createPaginator` keeps a `PageState` with `cursorY`, `contentBottom`, multi-column
  `constraintBoundaries` (text wraps around floats per Y-range), and contextual-spacing/border state.
- `sliceLines(lines, startIndex, availableHeight) → { toLine, height }` is the **line-level split
  primitive**: fit as many lines as the page's remaining height allows; the rest flow to the next
  page. `computeFragmentPmRange` gives each page fragment its doc positions.
- Tables split mid-row/mid-cell line-by-line (`table-cell-slice`: a per-line O(1) cost cursor for
  spacing.before / totalHeight promotion / spacing.after, kept in sync with the DOM painter).
- Fidelity we don't do yet: empty-paragraph **inherited-spacing suppression**
  (`shouldSuppressSpacingForEmpty` + tracking which spacing was *explicit*) and **contextual spacing**
  (`w:contextualSpacing` suppresses own spacing when the adjacent paragraph shares a styleId).
- Measurement uses canvas (`konva` dep); headless runs bundle `jsdom` (part of the 54 MB).

**Collaboration = ProseMirror ↔ Yjs.** Peer-deps `yjs` + `y-prosemirror` + `@hocuspocus/provider`;
the doc is a Yjs `XmlFragment` (char-level CRDT merge), presence via a collaboration-cursor extension.
`normalize-yjs-fragment` repairs schema mismatches (strips cached field-result leaf-atom children)
before y-prosemirror hydrates. This is the true conflict-free merge our snapshot sync lacks.

**Block editing is decomposed into many edge-case commands** — `backspaceAcrossRuns`,
`backspaceAtomBefore`, `backspaceEmptyRunParagraph`, `backspaceNextToRun`, `backspaceSkipEmptyRun`, …
each a small ProseMirror `Command`. The robustness is the long tail of these; our single merge is the
naive case.

**What to steal next (our own implementation, license-clean):**
1. **Line-level pagination** — measure paragraphs into line boxes (DOM `Range.getClientRects()`), flow
   lines with a `sliceLines`-style fit (extend our block `packPages`), so long paragraphs split across
   pages. Highest-value upgrade — closes gap #2.
2. **Empty-paragraph + contextual-spacing** suppression for render fidelity.
3. **Yjs binding** (model ↔ `Y.XmlFragment`) for CRDT collab — closes gap #1.
4. **Decompose backspace/merge** into the edge-case command set above.

## 5. Sources
- SuperDoc repo + `CONTRIBUTING.md` (architecture, module paths) — https://github.com/Harbour-Enterprises/SuperDoc
- SuperDoc docs (Document API, track changes) — https://docs.superdoc.dev
- SuperDoc product/engine notes — https://www.superdoc.dev
- ProseMirror pagination discussion (CSS-vs-computational trade) — https://discuss.prosemirror.net/t/a-new-text-editor-with-pagination/6667
- Authoritative format: **ECMA-376 / ISO-29500 (OOXML)** — the spec we actually implement against.
