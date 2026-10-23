import { PptxHandler } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { ViewerStateBag } from './create-viewer-state-types';
import CreateViewerStateHarness from './CreateViewerStateHarness.svelte';

/**
 * The find bar through `createViewerState`, so the viewer's own slide-change effect runs: it drops
 * the selection after every slide change, which used to wipe the selection Find had just made on
 * the match's slide. A slide change of the user's own must still drop the selection.
 */

let cleanup: (() => void) | undefined;
beforeEach(() => {
	localStorage.clear();
});
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

/** Two slides with one "quarter" match each. */
async function deck(): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
	data.slides.forEach((slide, index) => {
		const value = index ? 'quarter two' : 'quarter one';
		slide.elements = [
			{ id: `t${index}`, type: 'text', x: 67, y: 40, width: 400, height: 60, text: value },
		];
		slide.isDirty = true;
	});
	const bytes = await handler.save(data.slides);
	handler.dispose();
	return bytes;
}

async function loadedViewer(): Promise<ViewerStateBag> {
	let captured: ViewerStateBag | undefined;
	const instance = mount(CreateViewerStateHarness, {
		target: document.createElement('div'),
		props: {
			source: await deck(),
			editable: true,
			onready: (state: ViewerStateBag) => {
				captured = state;
			},
		},
	});
	cleanup = () => unmount(instance);
	const state = captured!;
	await vi.waitFor(
		() => {
			flushSync();
			expect(state.editor.slides).toHaveLength(2);
		},
		{ timeout: 15000 },
	);
	return state;
}

const idOf = (state: ViewerStateBag, slide: number, value: string) =>
	state.editor.slides[slide]?.elements.find((e) => (e as { text?: string }).text === value)?.id;
const selected = (state: ViewerStateBag) => [...state.editor.selection.ids];

describe('createViewerState find and replace', () => {
	it('next selects the match on another slide, and a slide change of its own still clears', async () => {
		const state = await loadedViewer();
		state.findReplace.toggle();
		state.findReplace.query = 'quarter';
		flushSync();
		expect(state.viewer.current).toBe(0);
		expect(selected(state)).toStrictEqual([idOf(state, 0, 'quarter one')]);

		state.findReplace.next();
		flushSync();
		expect(state.viewer.current).toBe(1);
		expect(selected(state)).toStrictEqual([idOf(state, 1, 'quarter two')]);

		state.viewer.goTo(0);
		flushSync();
		expect(selected(state)).toStrictEqual([]);
		state.editor.select(idOf(state, 0, 'quarter one')!);
		flushSync();
		state.viewer.goTo(1);
		flushSync();
		expect(selected(state)).toStrictEqual([]);
	});
});
