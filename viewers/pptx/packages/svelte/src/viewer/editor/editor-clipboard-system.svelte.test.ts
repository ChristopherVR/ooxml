import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { EditorState } from './editor-state.svelte';

const box = (id: string, text: string) =>
	({ id, type: 'text', x: 0, y: 0, width: 10, height: 10, text }) as unknown as PptxElement;

afterEach(() => {
	vi.unstubAllGlobals();
});

describe('Svelte copy writes the system clipboard', () => {
	it('keeps the internal payload and offers the selection text to other apps', async () => {
		const writeText = vi.fn(async () => undefined);
		vi.stubGlobal('ClipboardItem', undefined);
		vi.stubGlobal('navigator', { clipboard: { writeText } });
		const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
		editor.setSlides([
			{ id: 's1', elements: [box('a', 'Alpha'), box('b', 'Beta')] } as unknown as PptxSlide,
		]);
		editor.selection.setAll(['a', 'b']);
		editor.clipboardOps.copySelected();
		expect(editor.clipboard).not.toBeNull();
		await vi.waitFor(() => expect(writeText).toHaveBeenCalledWith('Alpha\nBeta'));
	});
});
