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
export { inspectorOptionLabel } from './inspector-option-label';
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
	sizePositionSections,
	ZOOM_OPTIONS_SECTION,
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
export {
	PICTURE_DEFAULT_COLOR_CHANGE,
	PICTURE_DEFAULT_COLOR_WASH,
	PICTURE_DEFAULT_DUOTONE,
	PICTURE_COLOR_WASH_SECTION,
	PICTURE_DUOTONE_SECTION,
	PICTURE_EFFECT_SECTIONS,
	PICTURE_RECOLOR_SECTION,
	pictureDuotonePresetPatch,
} from './picture-effect-fields';
export type { PictureEffectSection } from './picture-effect-fields';
export { PICTURE_ARTISTIC_SECTION, PICTURE_ARTISTIC_SECTIONS } from './picture-artistic-fields';
export type { PictureArtisticSection } from './picture-artistic-fields';
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
export {
	TABLE_CELL_MARGIN_SECTION,
	TABLE_CELL_SECTIONS,
	TABLE_CELL_TEXT_SECTION,
} from './table-cell-fields';
export type { TableCellSection, TableCellTarget } from './table-cell-fields';
export { SLIDE_BACKGROUND_SECTION, SLIDE_BACKGROUND_SECTIONS } from './slide-background-fields';
export type { SlideBackgroundSection, SlideBackgroundTarget } from './slide-background-fields';
export { MEDIA_PLAYBACK_SECTION, MEDIA_PLAYBACK_SECTIONS } from './media-playback-fields';
export type { MediaPlaybackSection } from './media-playback-fields';
export { MEDIA_TRIM_SECTION, MEDIA_TRIM_SECTIONS } from './media-trim-fields';
export type { MediaTrimSection, MediaTrimTarget } from './media-trim-fields';
export {
	OLE_OBJECT_NAME_SECTION,
	OLE_OBJECT_NAME_SECTIONS,
	resolveOleSummaryRows,
} from './ole-fields';
export type { OleSection, OleSummaryRow } from './ole-fields';
export {
	ACTION_SETTINGS_SECTION,
	ACTION_SETTINGS_SECTIONS,
	actionSettingsTargetOf,
	actionSettingsType,
} from './action-settings-fields';
export type {
	ActionSettingsChange,
	ActionSettingsSection,
	ActionSettingsTarget,
} from './action-settings-fields';
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
