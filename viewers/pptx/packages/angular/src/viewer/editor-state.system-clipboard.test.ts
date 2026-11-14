import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { EditorStateService } from './editor-state.service';

const box = (id: string, text: string) =>
	({ id, type: 'text', x: 0, y: 0, width: 10, height: 10, text }) as unknown as PptxElement;

afterEach(() => {
	vi.unstubAllGlobals();
});

describe('Angular copy writes the system clipboard', () => {
	it('keeps the internal snapshot and offers the selection text to other apps', async () => {
		const writeText = vi.fn(async () => undefined);
		vi.stubGlobal('ClipboardItem', undefined);
		vi.stubGlobal('navigator', { clipboard: { writeText } });
		const svc = new EditorStateService();
		svc.setSlides([
			{ id: 's1', elements: [box('a', 'Alpha'), box('b', 'Beta')] } as unknown as PptxSlide,
		]);
		svc.select(['a', 'b']);
		svc.copySelected(0);
		expect(svc.hasClipboard()).toBe(true);
		await vi.waitFor(() => expect(writeText).toHaveBeenCalledWith('Alpha\nBeta'));
	});

	it('still copies internally when the clipboard permission is denied', async () => {
		const writeText = vi.fn(async () => {
			throw Object.assign(new Error('denied'), { name: 'NotAllowedError' });
		});
		vi.stubGlobal('ClipboardItem', undefined);
		vi.stubGlobal('navigator', { clipboard: { writeText } });
		const svc = new EditorStateService();
		svc.setSlides([{ id: 's1', elements: [box('a', 'Alpha')] } as unknown as PptxSlide]);
		svc.select(['a']);
		svc.copySelected(0);
		await vi.waitFor(() => expect(writeText).toHaveBeenCalled());
		expect(svc.hasClipboard()).toBe(true);
	});
});
