/**
 * The Design galleries' hover preview and their effect and background commits over the editor
 * state. Shared (`createDeckStyleHost`) decides what a pick does to the slides, the template
 * elements and the theme; this reads and writes the editor, raises {@link EditorState.deckStyleHeld}
 * while a preview shows (autosave and collaboration then keep to the real document) and records
 * one undo step when a pick commits.
 */
import { createDeckStyleHost } from 'ooxml-ui/pptx';
import type { DeckStyleHandler, DeckStyleHost, DeckStyleState } from 'ooxml-ui/pptx';
import type { PptxTheme } from 'pptx-viewer-core';

import type { EditorState } from './editor-state.svelte';

export function createEditorDeckStyleHost(
	editor: EditorState,
	onthemechange?: (theme: PptxTheme | undefined) => void,
): DeckStyleHost {
	const write = (state: DeckStyleState): void => {
		editor.slides = state.slides;
		editor.templateElementsBySlideId = state.templateElementsBySlideId;
		editor.theme = state.theme;
	};
	return createDeckStyleHost(() => ({
		read: () => ({
			slides: editor.slides,
			templateElementsBySlideId: editor.templateElementsBySlideId,
			theme: editor.theme,
			themeColorMap: editor.themeColorMap,
		}),
		show: write,
		hold: (on) => {
			if (on) {
				editor.heldRenderedSlides = editor.renderedSlides;
			}
			editor.deckStyleHeld = on;
			if (!on) {
				editor.heldRenderedSlides = null;
			}
		},
		handler: () => editor.getHandler() as unknown as DeckStyleHandler | null,
		selectedSlideIds: () => {
			const slide = editor.slides[editor.currentSlideIndex];
			return slide ? [slide.id] : [];
		},
		apply: (state) => {
			editor.commitSlides(state.slides);
			editor.templateElementsBySlideId = state.templateElementsBySlideId;
			editor.theme = state.theme;
			onthemechange?.(state.theme);
		},
	}));
}
