// Compatibility exports: document operations live in ooxml-core.
export {
	findSlideElement,
	mapSlideElements,
	updateElement,
	patchElementGeometry,
	removeElement,
	duplicateElementOnSlide,
	cloneSlides,
	reorderElementOnSlide,
	appendElementOnSlide,
	updateSlideNotes,
	updateSlide,
	updateAllSlides,
} from 'ooxml-core/pptx/editor/render/editor-mutations';
export type { ElementBoxPatch } from 'ooxml-core/pptx/editor/render/editor-mutations';
