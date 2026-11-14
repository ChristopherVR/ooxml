import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { ref } from 'vue';

import { useElementClipboard } from './useElementClipboard';

const box = (id: string, text: string) =>
	({ id, type: 'text', x: 0, y: 0, width: 10, height: 10, text }) as unknown as PptxElement;

afterEach(() => {
	vi.unstubAllGlobals();
});

describe('Vue copy writes the system clipboard', () => {
	it('keeps the internal buffer and offers the selection text to other apps', async () => {
		const writeText = vi.fn(async () => undefined);
		vi.stubGlobal('ClipboardItem', undefined);
		vi.stubGlobal('navigator', { clipboard: { writeText } });
		const elements = new Map([
			['a', box('a', 'Alpha')],
			['b', box('b', 'Beta')],
		]);
		const clipboard = useElementClipboard({
			findSlideElement: (id) => elements.get(id),
			addElement: vi.fn(),
			removeElement: vi.fn(),
			selectedElementIds: ref(['a', 'b']),
			rasterize: vi.fn(async () => null),
		});
		clipboard.copyElement('a');
		expect(clipboard.hasClipboard.value).toBe(true);
		await vi.waitFor(() => expect(writeText).toHaveBeenCalledWith('Alpha\nBeta'));
	});
});
