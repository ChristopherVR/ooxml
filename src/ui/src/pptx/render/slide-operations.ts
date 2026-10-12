// Compatibility exports: document operations live in ooxml-core.
export {
	makeSlideId,
	createBlankSlide,
	captureSlideClipboard,
	createSlidePaste,
	insertSlidePaste,
	applySlidePasteTemplates,
	captureSlideClipboardTemplates,
} from 'ooxml-core/pptx/editor/render/slide-operations';
export type { SlidePasteRequest } from 'ooxml-core/pptx/editor/render/slide-operations';
