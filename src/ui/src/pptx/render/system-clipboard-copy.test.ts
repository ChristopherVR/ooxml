import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import {
	copyElementsToSystemClipboard,
	elementsClipboardText,
	writeSystemClipboard,
} from './system-clipboard-copy';

class FakeItem {
	static supports = (type: string) => type === 'image/png' || type === 'text/plain';
	constructor(readonly items: Record<string, Blob | PromiseLike<Blob>>) {}
}
const Item = FakeItem as unknown as Parameters<typeof writeSystemClipboard>[2];
const png = new Blob([new Uint8Array([137, 80, 78, 71])], { type: 'image/png' });

const textBox = (id: string, text: string) =>
	({ id, type: 'text', x: 0, y: 0, width: 10, height: 10, text }) as unknown as PptxElement;

describe('writeSystemClipboard', () => {
	it('writes the picture and the text together', async () => {
		const write = vi.fn(async () => undefined);
		const status = await writeSystemClipboard(
			{ text: 'Hi', image: Promise.resolve(png) },
			{ write },
			Item,
		);
		expect(status).toBe('image-and-text');
		const item = (write.mock.calls[0] as unknown as [FakeItem[]])[0][0] as FakeItem;
		expect(Object.keys(item.items)).toEqual(['image/png', 'text/plain']);
		expect(await item.items['image/png']).toBe(png);
	});

	it('falls back to text when rendering fails', async () => {
		const write = vi.fn(async (items: ClipboardItem[]) => {
			const item = items[0] as unknown as FakeItem;
			await Promise.all(Object.values(item.items));
		});
		const status = await writeSystemClipboard(
			{ text: 'Hi', image: Promise.resolve(null) },
			{ write },
			Item,
		);
		expect(status).toBe('text');
		expect(Object.keys((write.mock.calls[1]![0][0] as unknown as FakeItem).items)).toEqual([
			'text/plain',
		]);
	});

	it('reports a denied permission without throwing', async () => {
		const denied = Object.assign(new Error('denied'), { name: 'NotAllowedError' });
		const write = vi.fn(async () => {
			throw denied;
		});
		await expect(
			writeSystemClipboard({ text: 'Hi', image: Promise.resolve(png) }, { write }, Item),
		).resolves.toBe('failed');
		await expect(writeSystemClipboard({ text: 'Hi' }, undefined, Item)).resolves.toBe(
			'unavailable',
		);
	});

	it('uses writeText where ClipboardItem is missing', async () => {
		const writeText = vi.fn(async () => undefined);
		await expect(
			writeSystemClipboard({ text: 'Hi', image: Promise.resolve(png) }, { writeText }, undefined),
		).resolves.toBe('text');
		expect(writeText).toHaveBeenCalledWith('Hi');
	});
});

describe('copyElementsToSystemClipboard', () => {
	it('joins the copied shapes text and skips the picture without a rasteriser', async () => {
		expect(
			elementsClipboardText([textBox('a', 'One'), textBox('b', ''), textBox('c', 'Two')]),
		).toBe('One\nTwo');
		const writeText = vi.fn(async () => undefined);
		const status = await copyElementsToSystemClipboard({
			elements: [textBox('a', 'One')],
			target: { writeText },
		});
		expect(status).toBe('text');
		expect(writeText).toHaveBeenCalledWith('One');
	});
});
