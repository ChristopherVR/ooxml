// Compatibility exports: document operations live in ooxml-core.
export {
	makeSlideId,
	remapZoomTargets,
	createBlankSlide,
	captureSlideClipboard,
	createSlidePaste,
	insertSlidePaste,
	applySlidePasteTemplates,
	captureSlideClipboardTemplates,
} from 'ooxml-core/pptx/editor/render/slide-operations';
export type { SlidePasteRequest } from 'ooxml-core/pptx/editor/render/slide-operations';
