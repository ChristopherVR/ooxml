// Compatibility exports: document operations live in ooxml-core.
export {
	TABLE_STYLE_EDITOR_PARTS,
	isTableStylePartName,
	TABLE_STYLE_BORDER_SIDES,
	TABLE_STYLE_BORDER_SIDE_LABEL_KEYS,
	TABLE_STYLE_DASH_PRESETS,
} from 'ooxml-core/pptx/editor/render/table-style-editor-parts';
export type {
	TableStyleEditorPartId,
	TableStyleEditorPartDescriptor,
	TableStyleBorderSide,
} from 'ooxml-core/pptx/editor/render/table-style-editor-parts';
