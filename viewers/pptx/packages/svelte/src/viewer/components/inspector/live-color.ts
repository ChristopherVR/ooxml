/**
 * live-color.ts: the inspector's native colour pickers, live while the user
 * drags (the shared `createLiveColorInput` rule). The shades, and a drag put
 * back, go through `EditorState.previewElement`; the settled pick commits
 * through `patchSelected`.
 *
 * @module viewer/components/inspector/live-color
 */
import { createLiveColorInput, handleLiveColorEvent } from 'ooxml-ui/pptx';
import type { LiveColorInput } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import type { EditorState } from '../../editor/editor-state.svelte';

/**
 * The live colour controller for one element colour. `editor` and `element`
 * read the section's current props; `colorOf` and `patchOf` are the colour's
 * reader and its patch builder.
 */
export function liveElementColor(
	editor: () => EditorState,
	element: () => PptxElement,
	colorOf: (el: PptxElement) => string,
	patchOf: (el: PptxElement, hex: string) => Partial<PptxElement>,
): LiveColorInput {
	return createLiveColorInput<PptxElement, Partial<PptxElement>>({
		read: () => {
			const el = element();
			return { key: el.id, color: colorOf(el), value: el };
		},
		patch: patchOf,
		lookup: (key) => editor().elementById(key),
		preview: (patch, key) => editor().previewElement(key, patch),
		commit: (patch) => editor().patchSelected(patch),
		remember: (hex) => editor().recordRecentColor(hex),
	});
}

const EVENTS = ['input', 'change', 'blur'] as const;

/**
 * Svelte action for a native colour input: routes its `input`, `change` and
 * `blur` events to `control` (`use:liveColorEvents={control}`).
 */
export function liveColorEvents(
	node: HTMLInputElement,
	control: LiveColorInput,
): { update(next: LiveColorInput): void; destroy(): void } {
	let current = control;
	const handle = (event: Event): void => handleLiveColorEvent(current, event);
	for (const type of EVENTS) {
		node.addEventListener(type, handle);
	}
	return {
		update(next) {
			current = next;
		},
		destroy() {
			for (const type of EVENTS) {
				node.removeEventListener(type, handle);
			}
		},
	};
}
