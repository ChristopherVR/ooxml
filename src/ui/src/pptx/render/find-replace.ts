// Compatibility exports: document operations live in ooxml-core.
export {
	findInSlides,
	findReplacementAvailability,
	applyFindReplacements,
	replaceMatch,
	replaceInSlides,
} from 'ooxml-core/pptx/editor/render/find-replace';
export type {
	FindResult,
	FindOptions,
	ReplaceResult,
} from 'ooxml-core/pptx/editor/render/find-replace';
