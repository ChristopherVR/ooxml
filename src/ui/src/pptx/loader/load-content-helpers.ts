// Compatibility exports: document operations live in ooxml-core.
export {
	collectMediaElements,
	collectAnimationSoundPaths,
	collectImagePaths,
	collectTableCellImagePaths,
	applyTableCellImagePatches,
	buildInitialGuides,
} from 'ooxml-core/pptx/editor/loader/load-content-helpers';
export type {
	GuideEntry,
	ImagePathElement,
	TableCellImageRef,
} from 'ooxml-core/pptx/editor/loader/load-content-helpers';
