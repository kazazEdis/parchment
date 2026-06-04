# Parchment 📜

A headless, **MIT-licensed** `.docx` engine and interactive web editor for the browser **and** Node.

Parchment reads, renders, edits, and writes real Word documents with **byte-verbatim
preserve-and-patch fidelity** — it keeps every original part's bytes and rewrites only the nodes you
edit, so anything it doesn't model survives a round-trip untouched. It was built by studying
[SuperDoc](https://github.com/Harbour-Enterprises/SuperDoc)'s architecture and implementing against
the open **ECMA-376 / OOXML** spec — it contains **no SuperDoc code**.

## Why

| | SuperDoc | Parchment |
|---|---|---|
| Licence | AGPL-3.0 / paid | **MIT** |
| Footprint | ~54 MB package | small (jszip + yjs) |
| Runtime | browser + ProseMirror | **fully headless** — runs in Node / serverless too |
| Fidelity | re-emits via PM→OOXML | **byte-verbatim preserve-and-patch** |
| Tests | — | **300+ deterministic unit tests** |

## Features

- **Headless Document API** — open → query/edit → save, run-split-aware template fill, accept/reject changes
- **Interactive editor** (React) — inline selection + per-range formatting, undo/redo, paste-from-Word, block ops (split/merge/indent), find & replace
- **Track changes** + **generate redlines** from a word-level diff between two versions
- **Line-level pagination** (measure → page sheets, splits long paragraphs across pages)
- **CRDT collaboration** (Yjs over `document.xml`) + **presence cursors**
- **Comments** (read, author, sub-paragraph ranges) · **tables** (insert/delete rows+cols) · **images** (insert + render) · **hyperlinks** · **bookmarks**
- **Fields/TOC** · **footnotes** · **headers/footers** · **math** (OMML → MathML) · **encrypted-docx detection** · **AI-assist hook**

## Install

```bash
npm i github:kazazEdis/parchment
```

Parchment ships TypeScript source. In Next.js, add it to `transpilePackages`:

```ts
// next.config.ts
export default { transpilePackages: ["parchment"] };
```

## Usage — headless

```ts
import { openDocx, fillTemplate, save } from "parchment";

const doc = await openDocx(bytes);                 // Uint8Array | ArrayBuffer | Blob
const { doc: filled } = fillTemplate(doc, {        // run-split aware
  customer: "ACME d.o.o.",
  total: "1.234,56",
});
const out = await save(filled);                    // Uint8Array (.docx)
```

```ts
import { openDocx, redlineParagraph, patchParagraph } from "parchment";
// generate tracked-change redlines between two versions of a paragraph, etc.
```

## Usage — React

```tsx
import { DocxView, DocxEditor, PaginatedDocxView } from "parchment/react";
import { readDocx } from "parchment";

const pkg = await readDocx(bytes);
<DocxEditor initialPackage={pkg} collabChannel="my-doc" onAiRewrite={callClaude} />
```

## Architecture

The full design — OPC container, model, style-engine, preserve-and-patch writer, the SuperDoc
study, and the line-flow paginator — is in [ARCHITECTURE.md](./ARCHITECTURE.md).

## Develop

```bash
npm install
npm test          # vitest — 300+ unit tests
npm run typecheck # tsc --noEmit
```

## License

MIT © Edis Kazaz
