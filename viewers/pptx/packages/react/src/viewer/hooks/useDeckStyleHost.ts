import { createDeckStyleHost } from 'ooxml-ui/pptx';
import type { DeckStyleHandler, DeckStyleHost, DeckStyleState } from 'ooxml-ui/pptx';
import type { PptxElement, PptxHandler, PptxSlide, PptxTheme } from 'pptx-viewer-core';
/**
 * useDeckStyleHost: the Design galleries' hover preview and their commit, over this binding's
 * state. Shared (`createDeckStyleHost`) decides what a pick does to the slides, the master and
 * template elements and the theme; this hook only reads and writes the React state, holds the
 * undo tracker while a preview shows, and records one undo step when a pick commits.
 */
import { useMemo, useRef } from 'react';
import type { Dispatch, RefObject, SetStateAction } from 'react';

export interface UseDeckStyleHostInput {
	slides: PptxSlide[];
	templateElementsBySlideId: Record<string, PptxElement[]>;
	theme: PptxTheme | undefined;
	themeColorMap: Record<string, string> | undefined;
	handlerRef: RefObject<PptxHandler | null>;
	/** Slides a "selected slides" pick lands on (the active slide). */
	selectedSlideIds: readonly string[];
	setSlides: Dispatch<SetStateAction<PptxSlide[]>>;
	setTemplateElementsBySlideId: Dispatch<SetStateAction<Record<string, PptxElement[]>>>;
	setTheme: Dispatch<SetStateAction<PptxTheme | undefined>>;
	/** Pause and resume the undo tracker around a preview. */
	holdTracking: (held: boolean) => void;
	/** Open the undo tracker's gate so the committed change is recorded as one step. */
	bumpHistory: () => void;
	markDirty: () => void;
}

export function useDeckStyleHost(input: UseDeckStyleHostInput): DeckStyleHost {
	const latest = useRef(input);
	latest.current = input;
	return useMemo(
		() =>
			createDeckStyleHost(() => {
				const current = latest.current;
				const write = (state: DeckStyleState): void => {
					current.setSlides(state.slides);
					current.setTemplateElementsBySlideId(state.templateElementsBySlideId);
					current.setTheme(state.theme);
				};
				return {
					read: () => ({
						slides: current.slides,
						templateElementsBySlideId: current.templateElementsBySlideId,
						theme: current.theme,
						themeColorMap: current.themeColorMap,
					}),
					show: write,
					hold: current.holdTracking,
					handler: () => current.handlerRef.current as unknown as DeckStyleHandler | null,
					selectedSlideIds: () => current.selectedSlideIds,
					apply: (state) => {
						current.bumpHistory();
						write(state);
						current.markDirty();
					},
				};
			}),
		[],
	);
}
