/**
 * Load-pipeline helpers: moved to the framework-agnostic `pptx-viewer-shared`
 * package so the React, Vue, and Angular bindings share one copy.
 *
 * Re-exported here to keep existing import paths stable.
 */
export type {
	GuideEntry,
	ImagePathElement,
	TableCellImageRef,
	TableStyleImageRef,
	MediaArrayBufferSource,
	MediaSourceResolution,
} from 'ooxml-ui/pptx';
export {
	collectMediaElements,
	collectAnimationSoundPaths,
	collectImagePaths,
	collectTableCellImagePaths,
	applyTableCellImagePatches,
	collectTableStyleImagePaths,
	applyTableStyleImagePatches,
	buildInitialGuides,
	resolveMediaElementSource,
} from 'ooxml-ui/pptx';
