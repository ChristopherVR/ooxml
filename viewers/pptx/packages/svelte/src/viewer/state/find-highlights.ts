import type { ElementFindHighlights, SlideFindHighlights } from 'ooxml-ui/pptx';
import { getContext, setContext } from 'svelte';

/**
 * Svelte context for the Find bar's highlights on the editing canvas, by
 * element id (shared `findTextHighlightsByElement`), mirroring Vue's
 * `FindHighlightsKey`. Set by `SlideCanvas` only, so thumbnails, previews and
 * the show stage, which are not under it, never draw a mark. A getter, like the
 * slide-elements context next door, so a consumer's `$derived` stays reactive.
 */
export type FindHighlightsGetter = () => SlideFindHighlights | undefined;

/** Exported so tests can seed it through `mount(..., { context })`. */
export const FindHighlightsKey = Symbol('pptx-svelte-find-highlights');

/** Publish the canvas's find highlights (call at component initialisation). */
export function provideFindHighlights(getter: FindHighlightsGetter): void {
	setContext(FindHighlightsKey, getter);
}

/**
 * A getter of `elementId`'s highlights, or of nothing outside the editing
 * canvas. Call at initialisation and invoke inside a `$derived`.
 */
export function getElementFindHighlightsGetter(
	elementId: () => string,
): () => ElementFindHighlights | undefined {
	const source = getContext<FindHighlightsGetter | undefined>(FindHighlightsKey);
	return () => source?.()?.get(elementId());
}
