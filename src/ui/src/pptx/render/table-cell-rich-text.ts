// Compatibility exports: rich table cell text operations live in ooxml-core.
export {
	cellRunsFromSegments,
	setCellSegments,
	setTableElementCellSegments,
	tableCellSegments,
	withCellSegments,
} from 'ooxml-core/pptx/editor/render/table-cell-rich-text';
export {
	applyTableCellsTextStyle,
	withCellTextStyle,
} from 'ooxml-core/pptx/editor/render/table-cells-text-style';
export type {
	TableCellRef,
	TableCellTextStyleUpdate,
} from 'ooxml-core/pptx/editor/render/table-cells-text-style';
export { fillTableCells } from 'ooxml-core/pptx/editor/render/table-grid-fill';
export type { TableGridFill, TableGridWrite } from 'ooxml-core/pptx/editor/render/table-grid-fill';
export {
	editSegmentsText,
	segmentsPlainText,
} from 'ooxml-core/pptx/editor/render/segments-text-edit';
