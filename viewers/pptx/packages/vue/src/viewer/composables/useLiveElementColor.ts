/**
 * useLiveElementColor: an inspector panel's native colour picker, live while
 * the user drags (the shared `createLiveColorInput` rule). The shades, and a
 * drag put back, go through the injected preview channel
 * (`inspector-preview-context.ts`); the settled pick emits ONE `update`. Bind
 * the returned listeners with `v-on` on the `<input type="color">`.
 */
import { createLiveColorInput, handleLiveColorEvent } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import { injectInspectorPreview } from './inspector-preview-context';
import { injectRecentColors } from './recent-colors-context';

export interface UseLiveElementColorOptions {
	/** The panel's element now (read on every event). */
	element: () => PptxElement;
	/** The colour the picker edits on an element. */
	colorOf: (el: PptxElement) => string;
	/** The patch that sets `hex` on an element (a native pick clears any theme ref). */
	patch: (el: PptxElement, hex: string) => Partial<PptxElement>;
	/** Emit the settled pick's one `update`. */
	commit: (patch: Partial<PptxElement>) => void;
}

/** The `input` / `change` / `blur` listeners for one colour input. */
export type LiveColorListeners = Record<'input' | 'change' | 'blur', (event: Event) => void>;

export function useLiveElementColor(options: UseLiveElementColorOptions): LiveColorListeners {
	const preview = injectInspectorPreview();
	const recentColors = injectRecentColors();
	const control = createLiveColorInput<PptxElement, Partial<PptxElement>>({
		// The deck's own element, not the one last rendered: a drag rebases on it.
		read: () => {
			const shown = options.element();
			const el = preview?.element(shown.id) ?? shown;
			return { key: el.id, color: options.colorOf(el), value: el };
		},
		patch: options.patch,
		lookup: (key) => preview?.element(key),
		preview: (patch, key) => preview?.preview(patch, key),
		commit: (patch) => options.commit(patch),
		remember: (hex) => recentColors?.push(hex),
	});
	const handle = (event: Event): void => handleLiveColorEvent(control, event);
	return { input: handle, change: handle, blur: handle };
}
