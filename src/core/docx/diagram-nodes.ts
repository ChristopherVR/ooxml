// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The tree-order helpers moved to the format-neutral `diagram/edit/node-tree.ts`, which Excel
// shares; this module keeps the historical names.
export { diagramNodesInTreeOrder, treeOrderDiagramNodes } from '../diagram/index';
