/**
 * Descriptor-driven inspector sections: the field model every binding's one
 * generic field renderer draws, and the section descriptors written once.
 *
 * @module render/inspector
 */
export {
	coerceInspectorFieldValue,
	inspectorFieldPatch,
	inspectorTogglePatch,
	resolveInspectorSections,
} from './inspector-field-model';
export type {
	InspectorFieldControl,
	InspectorFieldDescriptor,
	InspectorFieldKind,
	InspectorFieldOption,
	InspectorFieldValue,
	InspectorFieldView,
	InspectorSectionDescriptor,
	InspectorSectionToggle,
	InspectorSectionView,
} from './inspector-field-model';
export { createInspectorColorFields } from './inspector-color-fields';
export type { InspectorColorFieldsOptions } from './inspector-color-fields';
export type {
	ShapeEffectField,
	ShapeEffectSection,
	ShapeEffectTarget,
} from './shape-effect-fields';
export { SHAPE_BEVEL_OPTIONS } from './shape-effect-fields-more';
export { SHAPE_EFFECT_SECTIONS } from './shape-effect-sections';
export { SHAPE_LINE_SECTION, SHAPE_LINE_SECTIONS } from './shape-line-fields';
export {
	DEFAULT_PATTERN_BACKGROUND,
	DEFAULT_PATTERN_PRESET,
	isUsableFillColor,
	SHAPE_FILL_MODE_OPTIONS,
	SHAPE_FILL_MODE_SECTION,
	SHAPE_FILL_MODE_SECTIONS,
	SHAPE_FILL_OPTIONS_SECTION,
	SHAPE_FILL_OPTIONS_SECTIONS,
	shapeFillColorOf,
	shapeFillModeChange,
	shapeFillModeOf,
} from './shape-fill-fields';
export type { ShapeFillMode, ShapeFillSection, ShapeFillTarget } from './shape-fill-fields';
export {
	addShapeFillStop,
	editShapeFillStop,
	removeShapeFillStop,
	shapeFillStopsOf,
} from './shape-fill-stops';
export type { ShapeFillStop } from './shape-fill-stops';
export { resolveShapeFillView, shapeFillPatch } from './shape-fill-actions';
export type {
	ShapeFillAction,
	ShapeFillPatternSwatch,
	ShapeFillStopView,
	ShapeFillView,
} from './shape-fill-actions';

export type { ShapeLineSection } from './shape-line-fields';
export {
	ELEMENT_GEOMETRY_SECTION,
	ELEMENT_GEOMETRY_SECTIONS,
	MIN_ELEMENT_SIZE,
} from './element-geometry-fields';
export type {
	ElementGeometryPatch,
	ElementGeometrySection,
	ElementGeometryTarget,
} from './element-geometry-fields';
export { TEXT_BOX_SECTION, TEXT_BOX_SECTIONS } from './text-box-fields';
export { ELEMENT_ALT_TEXT_SECTION, ELEMENT_ALT_TEXT_SECTIONS } from './element-alt-text-fields';
export type { ElementAltTextSection } from './element-alt-text-fields';
export {
	MAX_PICTURE_CROP_PERCENT,
	PICTURE_ADJUSTMENT_SECTION,
	PICTURE_ADJUSTMENT_SECTIONS,
	PICTURE_CROP_SECTION,
	PICTURE_CROP_SECTIONS,
} from './picture-fields';
export type { PictureSection } from './picture-fields';
export type { TextBoxSection, TextBoxTarget } from './text-box-fields';
export {
	CHARACTER_SPACING_SECTION,
	CHARACTER_SPACING_SECTIONS,
	PARAGRAPH_SPACING_SECTION,
	PARAGRAPH_SPACING_SECTIONS,
} from './paragraph-fields';
export type { ParagraphFieldsSection, ParagraphFieldsTarget } from './paragraph-fields';
export { TABLE_OPTIONS_SECTION, TABLE_OPTIONS_SECTIONS } from './table-option-fields';
export type { TableOptionsSection, TableOptionsTarget } from './table-option-fields';
export { TEXT_EFFECT_SECTIONS } from './text-effect-fields';
export type { TextEffectSection, TextEffectTarget } from './text-effect-fields';
export { FORMAT_PANEL_SECTIONS, formatPanelIn } from './format-panel-sections';
export type { FormatPanelId } from './format-panel-sections';
export {
	FORMAT_SECTION_ATTRIBUTE,
	FORMAT_SECTION_IDS,
	formatElementCategory,
	hasFormatShapeEffect,
	isLineFormatElement,
	resolveFormatCategory,
	resolveFormatSections,
} from './format-sections';
export type {
	FormatSectionId,
	FormatSectionOptions,
	FormatSectionView,
	FormatSelectionCategory,
} from './format-sections';
export {
	formatMixedValue,
	formatSelectionElements,
	isFormatSectionMixed,
} from './format-section-mixed';
export type { FormatFieldValue } from './format-section-mixed';
export {
	findFormatSection,
	formatSectionForAnchor,
	revealFormatSection,
} from './format-section-reveal';
