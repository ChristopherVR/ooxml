import { createDeckStyleHost } from 'ooxml-ui/pptx';
import type { DeckStyleHandler, DeckStyleHost, DeckStyleState } from 'ooxml-ui/pptx';
import type { PptxElement, PptxHandler, PptxSlide, PptxTheme } from 'pptx-viewer-core';
/**
 * useDeckStyleHost: the Design galleries' hover preview and commit over the Vue refs. Shared
 * (`createDeckStyleHost`) decides what a pick does to the slides, template elements and theme;
 * this only reads and writes the refs, raises `held` while a preview shows (autosave and
 * collaboration stand still), and records one undo step when a pick commits.
 */
import type { Ref, ShallowRef } from 'vue';

export interface UseDeckStyleHostInput {
	slides: Ref<PptxSlide[]>;
	templateElementsBySlideId: ShallowRef<Record<string, PptxElement[]>>;
	pptxTheme: Ref<PptxTheme | undefined>;
	themeColorMap: Ref<Record<string, string> | undefined>;
	handler: Ref<PptxHandler | null | undefined>;
	/** Slides a "selected slides" pick lands on (the active slide). */
	selectedSlideIds: () => readonly string[];
	/** Snapshot the current state onto the undo stack, right before a commit. */
	pushHistory: (label?: string) => void;
	/**
	 * Raised while a hover preview shows over the real slides. The same ref is given to autosave
	 * and collaboration, whose synchronous watchers must see it before the preview's first write.
	 */
	held: Ref<boolean>;
}

export function useDeckStyleHost(input: UseDeckStyleHostInput): DeckStyleHost {
	const write = (state: DeckStyleState): void => {
		input.slides.value = state.slides;
		input.templateElementsBySlideId.value = state.templateElementsBySlideId;
		input.pptxTheme.value = state.theme;
		input.themeColorMap.value = state.themeColorMap;
	};
	return createDeckStyleHost(() => ({
		read: () => ({
			slides: input.slides.value,
			templateElementsBySlideId: input.templateElementsBySlideId.value,
			theme: input.pptxTheme.value,
			themeColorMap: input.themeColorMap.value,
		}),
		show: write,
		hold: (on) => {
			input.held.value = on;
		},
		handler: () => input.handler.value as unknown as DeckStyleHandler | null | undefined,
		selectedSlideIds: input.selectedSlideIds,
		apply: (state) => {
			input.pushHistory();
			write(state);
		},
	}));
}
