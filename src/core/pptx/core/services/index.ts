/**
 * Barrel export for the PPTX services layer.
 *
 * Re-exports all public service classes, interfaces, and supporting types
 * used by the core PPTX handler for slide loading, animation parsing/writing,
 * transitions, compatibility inspection, XML lookup, document properties,
 * and template backgrounds.
 *
 * @module services
 */

export {
	PptxCompatibilityService,
	type CompatibilityWarningInput,
	type IPptxCompatibilityService,
} from './PptxCompatibilityService';
export { PptxDocumentPropertiesUpdater } from './PptxDocumentPropertiesUpdater';
export {
	PptxTemplateBackgroundService,
	type IPptxTemplateBackgroundService,
	type PptxTemplateBackgroundState,
} from './PptxTemplateBackgroundService';
export {
	PptxSlideLoaderService,
	type IPptxSlideLoaderService,
	type PptxMediaTimingEntry,
	type PptxMediaTimingMap,
	type PptxSlideLoaderParams,
	type PptxSlideLoaderThemeOverride,
	type PptxSlideNotesResult,
} from './PptxSlideLoaderService';
export { PptxXmlLookupService, type IPptxXmlLookupService } from './PptxXmlLookupService';
export {
	PptxEditorAnimationService,
	type IPptxEditorAnimationService,
	type PptxEditorAnimationServiceOptions,
} from './PptxEditorAnimationService';
export { reconcileAnimationTargets } from './animation-target-reconcile';
export { remapEditorAnimationsToShapeIds } from './animation-shape-id-assign';
export { resolveConnectorShapeIds, applyConnectorShapeIds } from './connector-shape-id-assign';
export { mergeNativeSoundIntoEditorAnimations } from './animation-sound-merge';
export {
	PptxNativeAnimationService,
	type IPptxNativeAnimationService,
} from './PptxNativeAnimationService';
export {
	PptxSlideTransitionService,
	type IPptxSlideTransitionService,
	type PptxSlideTransitionServiceOptions,
} from './PptxSlideTransitionService';
export {
	PptxAnimationWriteService,
	type IPptxAnimationWriteService,
} from './PptxAnimationWriteService';
export {
	PRESET_TO_OOXML,
	OOXML_TO_PRESET_ENTR,
	OOXML_TO_PRESET_EXIT,
	OOXML_TO_PRESET_EMPH,
	ooxmlToPresetName,
	resolveOoxmlPresetMapping,
	timingCurveToAccelDecel,
} from './animation-write-mappings';
export { resolveDirectionSubtype } from './animation-timing-surgical-behaviors';
export {
	capturedPresetIds,
	capturedSubtypes,
	getCapturedPreset,
	getCapturedTree,
	type CapturedBehaviorNode,
	type CapturedPreset,
} from './animation-behavior-captured';
export { capturedPresetNativeAnimation } from './animation-behavior-native';
export { buildSingleEffectNode, applySoundToEffectCTn } from './animation-write-node-builders';
export { applyAfterAnimationBehavior } from './animation-after-effect-write';
export {
	animationFor,
	hasAnimation,
	upsert,
	DEFAULT_ANIMATION_DURATION_MS,
	DEFAULT_ANIMATION_TRIGGER,
} from './editor-animation-entry';
export {
	AFTER_ANIMATION_VALUES,
	DEFAULT_AFTER_ANIMATION_DIM_COLOR,
	setAfterAnimation,
	setAfterAnimationColor,
	getAfterAnimation,
	getAfterAnimationState,
	isNativeAfterAnimationOnly,
	mergeNativeAfterAnimationIntoEditorAnimations,
} from './editor-animation-after-effect';
export { applyAfterAnimationFromEditorList } from './editor-animation-after-effect-precedence';
export { extractAfterAnimationFromSubTnLst } from './native-animation-after-effect';
export {
	parseCondition,
	parseConditionList,
	serializeCondition,
	serializeConditionList,
} from './native-animation-helpers';
export {
	parseTimeTargetElement,
	serializeTimeTargetElement,
	extractGraphicBuilds,
	serializeGraphicBuild,
	type PptxGraphicBuildEntry,
} from './animation-target-build-helpers';
export { parseAnimEffectFilter, parseFilterToken } from './animation-effect-filter-parsing';
export { resolveOpacitySet, type PptxOpacitySet } from './animation-opacity-set';
