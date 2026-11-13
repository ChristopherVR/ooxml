import type { ElementFindHighlights, SlideFindHighlights } from 'ooxml-ui/pptx';
import type { InjectionKey } from 'vue';
import { inject, provide } from 'vue';

/**
 * The Find bar's highlights on the editing canvas, by element id (shared
 * `findTextHighlightsByElement`). Provided by `SlideCanvas` only, so thumbnails,
 * previews and the show stage, which are not under it, never draw a mark.
 * Like the slide-elements injection next door, a getter keeps the consumer's
 * `computed` reactive without a prop on every element renderer.
 */
export const FindHighlightsKey: InjectionKey<() => SlideFindHighlights | undefined> = Symbol(
	'pptx-vue-find-highlights',
);

/** Publish the canvas's find highlights (call from `setup`). */
export function provideFindHighlights(source: () => SlideFindHighlights | undefined): void {
	provide(FindHighlightsKey, source);
}

/**
 * A getter of this element's highlights, or of nothing outside the editing
 * canvas (call from `setup`, read inside a `computed`).
 */
export function injectElementFindHighlights(
	elementId: () => string,
): () => ElementFindHighlights | undefined {
	const source = inject(FindHighlightsKey, undefined);
	return () => source?.()?.get(elementId());
}
