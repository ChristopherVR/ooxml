import { applyPresentationInk, isPresentationAudience } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { ref } from 'vue';
import type { Ref } from 'vue';

import type { SlideAnnotationMap } from './usePresentationAnnotations';

export interface UsePresentationModeWiringInput {
	slides: Ref<PptxSlide[]>;
	activeSlideIndex: Ref<number>;
	pushHistory: () => void;
}

export interface UsePresentationModeWiringResult {
	presenting: Ref<boolean>;
	startPresenting: () => void;
	onPresentClose: (payload?: { annotations: SlideAnnotationMap }) => void;
	onPresentSlideChange: (index: number) => void;
}

/**
 * usePresentationModeWiring: the slideshow presentation-mode lifecycle,
 * including persisting kept ink annotations (drawn during presentation) back
 * onto their slides as real `ink` elements on close. Extracted verbatim from
 * `PowerPointViewer.vue`.
 */
export function usePresentationModeWiring(
	input: UsePresentationModeWiringInput,
): UsePresentationModeWiringResult {
	const { slides, activeSlideIndex, pushHistory } = input;

	const presenting = ref(
		typeof window !== 'undefined' && isPresentationAudience(window.location.hash),
	);
	function startPresenting(): void {
		presenting.value = true;
	}
	function onPresentClose(payload?: { annotations: SlideAnnotationMap }): void {
		presenting.value = false;
		const map = payload?.annotations;
		if (!map || map.size === 0) {
			return;
		}
		// Kept ink becomes `ink` elements on the slides it was drawn on, committed
		// as one history-tracked change so the whole batch undoes together.
		const nextSlides = applyPresentationInk(slides.value, [...map.values()].flat());
		if (nextSlides !== slides.value) {
			pushHistory();
			slides.value = [...nextSlides];
		}
	}
	function onPresentSlideChange(index: number): void {
		activeSlideIndex.value = index;
	}

	return { presenting, startPresenting, onPresentClose, onPresentSlideChange };
}
