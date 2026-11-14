import { afterEach, describe, expect, it, vi } from 'vitest';
import {
	PptxHandler,
	deserializeSlidePackageClipboard,
	SLIDE_PACKAGE_CLIPBOARD_MIME,
} from 'ooxml-core/pptx';
import { createSlidePackageClipboard } from './slide-package-clipboard';
import type { SlidePackageClipboardSnapshot } from './slide-package-clipboard';

class FakeItem {
	static supports = () => true;
	constructor(readonly blobs: Record<string, Blob | Promise<Blob>>) {}
	get types() {
		return Object.keys(this.blobs);
	}
	async getType(type: string) {
		return this.blobs[type]!;
	}
}
function transport() {
	let items: FakeItem[] = [];
	const clipboard = {
		write: vi.fn(async (value: FakeItem[]) => {
			await Promise.all(value.flatMap((item) => Object.values(item.blobs)));
			items = value;
		}),
		read: vi.fn(async () => items),
		readText: vi.fn(async () => 'foreign clipboard text'),
	};
	vi.stubGlobal('ClipboardItem', FakeItem);
	vi.stubGlobal('navigator', { clipboard });
	return clipboard;
}
async function deck(label: string): Promise<SlidePackageClipboardSnapshot> {
	const { handler, data, createSlide } = await PptxHandler.createBlank();
	data.slides = [createSlide('Blank').addText(label).build()];
	return { handler, slides: data.slides, templateElementsBySlideId: {}, editable: true };
}
afterEach(() => {
	vi.unstubAllGlobals();
	vi.restoreAllMocks();
});

describe('slide package system clipboard controller', () => {
	it('copies a readable preview and isolated package, pastes into another deck and retains old resource ownership', async () => {
		const clipboard = transport(),
			source = await deck('Source'),
			destination = await deck('Destination');
		let current = source;
		const commit = vi.fn((result) => {
			current = {
				...current,
				handler: result.handler,
				slides: result.slides,
				templateElementsBySlideId: result.templateElementsBySlideId,
			};
		});
		const controller = createSlidePackageClipboard({ getSnapshot: () => current, commit });
		const disposed = vi.spyOn(destination.handler, 'dispose');
		await controller.copy([0]);
		const items = await clipboard.read();
		expect(await (await items[0]!.getType('text/plain')).text()).toBe('Source');
		const packageText = await (await items[0]!.getType(SLIDE_PACKAGE_CLIPBOARD_MIME)).text();
		expect(deserializeSlidePackageClipboard(packageText)?.byteLength).toBeGreaterThan(0);
		current = destination;
		await controller.paste(1);
		expect(commit).toHaveBeenCalledOnce();
		expect(current.slides).toHaveLength(2);
		expect(current.slides[1]!.elements.find((el) => el.type === 'text')?.text).toBe('Source');
		expect(disposed).not.toHaveBeenCalled();
		controller.dispose();
		expect(disposed).toHaveBeenCalledOnce();
		current.handler.dispose();
		source.handler.dispose();
	});
	it('adds the slide picture when it renders and keeps the package when it does not', async () => {
		const clipboard = transport(),
			current = await deck('Source');
		const png = new Blob([new Uint8Array([137, 80])], { type: 'image/png' });
		const renderSlidePng = vi.fn(async () => png);
		const controller = createSlidePackageClipboard({
			getSnapshot: () => current,
			commit: vi.fn(),
			renderSlidePng,
		});
		await controller.copy([0]);
		expect(renderSlidePng).toHaveBeenCalledWith(0);
		expect((await clipboard.read())[0]!.types).toContain('image/png');
		renderSlidePng.mockResolvedValueOnce(null as unknown as Blob);
		await controller.copy([0]);
		const types = (await clipboard.read())[0]!.types;
		expect(types).toContain(SLIDE_PACKAGE_CLIPBOARD_MIME);
		expect(types).not.toContain('image/png');
		controller.dispose();
		current.handler.dispose();
	});
	it('reports clipboard permission failures and refuses foreign text without local fallback', async () => {
		const clipboard = transport(),
			current = await deck('Source'),
			commit = vi.fn();
		const controller = createSlidePackageClipboard({ getSnapshot: () => current, commit });
		clipboard.write.mockRejectedValueOnce(new Error('Clipboard permission denied'));
		await expect(controller.copy([0])).rejects.toThrow('Clipboard permission denied');
		await expect(controller.paste(0)).rejects.toThrow('does not contain copied');
		expect(commit).not.toHaveBeenCalled();
		controller.dispose();
		current.handler.dispose();
	});
	it('rejects an asynchronous paste when the editable document changes', async () => {
		const clipboard = transport(),
			current = await deck('Source'),
			commit = vi.fn();
		let snapshot = current;
		const controller = createSlidePackageClipboard({ getSnapshot: () => snapshot, commit });
		await controller.copy([0]);
		const originalRead = clipboard.read.getMockImplementation()!;
		clipboard.read.mockImplementationOnce(async () => {
			snapshot = { ...current, slides: [...current.slides] };
			return originalRead();
		});
		await expect(controller.paste(1)).rejects.toThrow('Presentation changed');
		expect(commit).not.toHaveBeenCalled();
		controller.dispose();
		current.handler.dispose();
	});
});
