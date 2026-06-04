// Public API for the in-house .docx editor — the headless engine. See ARCHITECTURE.md for the
// design (OPC container → model → style-engine → preserve-and-patch writer) and the SuperDoc study.
//
// The React renderer is intentionally NOT re-exported here so headless consumers (Node, tests,
// server code) don't pull in React. Import it directly: `import { DocxView } from ".../DocxView"`.
export * from "./opc";        // zip ↔ parts (read/write a .docx package)
export * from "./xml";        // OOXML string scanner + surgical patcher
export * from "./units";      // twip / half-pt / EMU / pct conversions
export * from "./props";      // rPr / pPr → typed property bags + cascade merge
export * from "./styles";     // styles.xml + the property cascade
export * from "./numbering";  // numbering.xml + marker formatting
export * from "./model";      // document.xml → editable tree
export * from "./resolve";    // style-engine: effective props + list-number counter pass
export * from "./serialize";  // model edits → bytes (preserve-and-patch)
export * from "./cssMap";     // effective props → CSS
export * from "./table";      // gridSpan / vMerge → render grid
export * from "./images";     // blip relationship → data URL
export * from "./edit";       // pure editing operations
export * from "./tableEdit";  // table structure (insert/delete rows + columns)
export * from "./opcParts";   // relationship + content-type registration
export * from "./imageInsert"; // insert an image (media part + drawing run)
export * from "./linkEdit";   // apply a hyperlink to a selection
export * from "./diff";       // word-level diff → tracked-change redlines
export * from "./doc";        // headless Document API (open → query/edit → save)
export * from "./paginate";   // computed pagination (block packing)
export * from "./comments";   // comments.xml: parse, locate, author
export * from "./fields";     // field instructions + TOC generation
export * from "./footnotes";  // footnotes.xml parse + refs
export * from "./headerFooter"; // resolve header/footer part XML
export * from "./bookmarks";  // bookmark parse + insert
export * from "./encrypted";  // encrypted-docx detection
export * from "./math";       // OMML → MathML
export * from "./collab";     // collaboration transport (BroadcastChannel / pluggable)
export * from "./crdt";       // CRDT collaboration (Yjs over document.xml)
