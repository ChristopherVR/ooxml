import type { Ref } from 'vue';
import { watch } from 'vue';

/**
 * The viewer drops the selection when the active slide changes, and keeps it only when the change
 * came with its own selection: Find goes to a match's slide and selects the match there in one
 * step. {@link UseSlideChangeSelectionResult.showSlideSelection} makes that step and notes the
 * slide; the slide-change watcher keeps the selection once when the change lands on the noted
 * slide, and drops it as before for every other slide change (the slides pane, Page Down, the
 * host's `setActiveSlideIndex`). The same rule as React's `useSlideChangeSelection` and Svelte's
 * `ViewerState.goToWithSelection`.
 */
export interface UseSlideChangeSelectionInput {
	activeSlideIndex: Ref<number>;
	/** Go to a slide (clamped to the deck). */
	goTo: (index: number) => void;
	selectedElementIds: Ref<string[]>;
}

export interface UseSlideChangeSelectionResult {
	/** Go to `slideIndex` and select `ids` there, across the slide change. */
	showSlideSelection: (slideIndex: number, ids: string[]) => void;
}

export function useSlideChangeSelection({
	activeSlideIndex,
	goTo,
	selectedElementIds,
}: UseSlideChangeSelectionInput): UseSlideChangeSelectionResult {
	/** The slide the last `showSlideSelection` selected on, until the next slide change. */
	let selectionSlide: number | null = null;
	watch(activeSlideIndex, (index) => {
		const kept = selectionSlide === index;
		selectionSlide = null;
		if (!kept) {
			selectedElementIds.value = [];
		}
	});
	return {
		showSlideSelection: (slideIndex, ids) => {
			goTo(slideIndex);
			selectionSlide = activeSlideIndex.value;
			selectedElementIds.value = ids;
		},
	};
}
