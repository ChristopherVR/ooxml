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
export type {
	ShapeEffectField,
	ShapeEffectSection,
	ShapeEffectTarget,
} from './shape-effect-fields';
export { SHAPE_BEVEL_OPTIONS } from './shape-effect-fields-more';
export { SHAPE_EFFECT_SECTIONS } from './shape-effect-sections';
export { SHAPE_LINE_SECTION, SHAPE_LINE_SECTIONS } from './shape-line-fields';
export type { ShapeLineSection } from './shape-line-fields';
