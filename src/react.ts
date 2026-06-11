// React UI entry (import from "parchment/react"). Kept separate from the headless engine ("parchment")
// so Node/server consumers don't pull in React.
export { DocxView } from "./DocxView";
export { DocxEditor, type EditorToken, type EditorLabels } from "./DocxEditor";
export { PaginatedDocxView } from "./PaginatedDocxView";
