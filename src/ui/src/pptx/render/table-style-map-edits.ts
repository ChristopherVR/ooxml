// Compatibility exports: document operations live in ooxml-core.
export {
	applyTableStyleMapChange,
	applyTableStyleDelete,
	tableStyleSaveOptions,
	tableStyleAssignmentUpdate,
} from 'ooxml-core/pptx/editor/render/table-style-map-edits';
export type {
	TableStyleMapEditState,
	TableStyleMapEditResult,
	TableStyleSaveOptionsState,
	TableStyleSaveOptions,
} from 'ooxml-core/pptx/editor/render/table-style-map-edits';
