// @vitest-environment happy-dom
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { EditorState } from '../editor/editor-state.svelte';
import { CompareController } from './compare-controller.svelte';

/** Answer the next file picker with a deck and capture the options Compare loads it with. */
function pickDeckAndCaptureLoad() {
	vi.spyOn(HTMLInputElement.prototype, 'click').mockImplementation(function (
		this: HTMLInputElement,
	) {
		Object.defineProperty(this, 'files', { value: [new File(['pptx'], 'other.pptx')] });
		void this.onchange?.(new Event('change'));
	});
	return vi.spyOn(PptxHandler.prototype, 'load').mockResolvedValue({ slides: [] } as never);
}

const editor = { renderedSlides: [] } as unknown as EditorState;

describe('CompareController linked content', () => {
	afterEach(() => vi.restoreAllMocks());

	it('loads the incoming deck with the resolved external-content switch', async () => {
		const load = pickDeckAndCaptureLoad();
		await new CompareController(editor, () => true).chooseFile();
		await vi.waitFor(() => expect(load).toHaveBeenCalled());
		expect(load.mock.calls[0][1]).toStrictEqual({ allowExternalImages: true });
	});

	it('blocks linked content when no switch is supplied', async () => {
		const load = pickDeckAndCaptureLoad();
		await new CompareController(editor).chooseFile();
		await vi.waitFor(() => expect(load).toHaveBeenCalled());
		expect(load.mock.calls[0][1]).toStrictEqual({ allowExternalImages: false });
	});
});
