/**
 * Injection point for the inspector's live preview channel: merge a patch into
 * an element WITHOUT an undo step (a shade of a colour picker drag, which
 * shared `createLiveColorInput` puts back before its one committed `update`),
 * and read an element as the deck holds it now, which the drag rebases on.
 * Provided once at the viewer root next to `RecentColorsKey`, so every colour
 * picker deep in the inspector tree reaches it without threading a second event
 * through each panel.
 */
import type { PptxElement } from 'pptx-viewer-core';
import type { InjectionKey } from 'vue';
import { inject } from 'vue';

export interface InspectorPreview {
	/**
	 * Merge `patch` into element `id` (the inspector's element, or one a drag has
	 * left) without an undo step.
	 */
	preview(patch: Partial<PptxElement>, id: string): void;
	/** Element `id` as the deck holds it now, not as the panels were last rendered. */
	element(id: string): PptxElement | undefined;
}

export const InspectorPreviewKey: InjectionKey<InspectorPreview> = Symbol(
	'pptx-vue-inspector-preview',
);

/**
 * The injected preview channel, or `undefined` outside a `PowerPointViewer`
 * (an isolated panel): a colour drag then shows nothing until it settles, and
 * still commits once.
 */
export function injectInspectorPreview(): InspectorPreview | undefined {
	return inject(InspectorPreviewKey, undefined);
}
