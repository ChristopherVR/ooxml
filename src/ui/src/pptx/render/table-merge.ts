// Compatibility exports: document operations live in ooxml-core.
export {
	computeBoundingRect,
	expandRectForExistingMerges,
	canMergeCells,
	canSplitCell,
	mergeCells,
	splitCell,
	computeSelectionRect,
	rectToCells,
	isCellInRect,
} from 'ooxml-core/pptx/editor/render/table-merge';
export type { CellCoord, CellRect } from 'ooxml-core/pptx/editor/render/table-merge';
