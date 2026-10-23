/**
 * Injection point for the inspector's live preview channel: merge a patch into
 * the inspector's element WITHOUT an undo step (a shade of a colour picker
 * drag, which shared `createLiveColorInput` puts back before its one committed
 * `update`). Provided once at the viewer root next to `RecentColorsKey`, so
 * every colour picker deep in the inspector tree reaches it without threading
 * a second event through each panel.
 */
import type { PptxElement } from 'pptx-viewer-core';
import type { InjectionKey } from 'vue';
import { inject } from 'vue';

export type InspectorPreview = (patch: Partial<PptxElement>) => void;

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
