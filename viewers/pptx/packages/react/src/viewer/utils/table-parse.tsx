// Barrel re-export: table style context and the raw-XML edit operations.
// Rendering reads the structured `tableData` (see `table-render.tsx`); the raw XML
// is only written back to, by the edit handlers.

export { type TableStyleContext, getTableCellBandStyle } from './table-band-style';

export {
	updateCellTextInRawXml,
	updateCellTextStyleInRawXml,
	updateMergeAttrsInRawXml,
	rebuildTableStructureInRawXml,
} from './table-xml-ops';
