/**
 * Options and callbacks the vanilla viewer gained to match the other four bindings
 * (`authorName`, `filePath`, `onOpenFile`, `onContentChange`, `onModeChange`,
 * `onSlideCountChange`), driven through the real viewer.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { afterEach, describe, expect, it, vi } from 'vitest';

import { createPptxViewer } from './PptxViewer';
import type { PptxViewerInstance, PptxViewerOptions } from './types';

const deck = new Uint8Array(
	readFileSync(resolve(process.cwd(), '../../../../e2e/pptx/fixtures/slide-jump-links.pptx')),
);

let viewer: PptxViewerInstance | undefined;
afterEach(() => {
	viewer?.destroy();
	viewer = undefined;
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

async function mountLoaded(options: PptxViewerOptions): Promise<PptxViewerInstance> {
	const container = document.createElement('div');
	document.body.append(container);
	viewer = createPptxViewer(container, options);
	await viewer.loadFile(deck);
	return viewer;
}

describe('vanilla callbacks shared with the other bindings', () => {
	it('reports the slide count when a slide is added or the deck is replaced', async () => {
		const onSlideCountChange = vi.fn();
		const instance = await mountLoaded({ editable: true, onSlideCountChange });
		const loaded = instance.getSlideCount();
		expect(onSlideCountChange).toHaveBeenLastCalledWith(loaded);

		instance.addSlide();

		expect(onSlideCountChange).toHaveBeenLastCalledWith(loaded + 1);
	});

	it('reports the viewer mode as editing is toggled', async () => {
		const onModeChange = vi.fn();
		const instance = await mountLoaded({ onModeChange });

		instance.setEditable(true);
		instance.setEditable(false);

		expect(onModeChange.mock.calls.map((call) => call[0])).toStrictEqual(['edit', 'preview']);
	});

	it('serialises the deck for onContentChange after a committed edit', async () => {
		const onContentChange = vi.fn();
		const instance = await mountLoaded({ editable: true, onContentChange });
		expect(onContentChange).not.toHaveBeenCalled();

		instance.addSlide();
		// ZIP serialization competes with the other viewer suites on hosted runners.
		await vi.waitFor(() => expect(onContentChange).toHaveBeenCalled(), { timeout: 10_000 });

		const bytes = onContentChange.mock.calls.at(-1)?.[0] as Uint8Array;
		expect(bytes).toBeInstanceOf(Uint8Array);
		// A .pptx is a ZIP package.
		expect(Array.from(bytes.slice(0, 2))).toStrictEqual([0x50, 0x4b]);
	});

	it('hands File > Open to the host when onOpenFile is supplied', async () => {
		const onOpenFile = vi.fn();
		const instance = await mountLoaded({ onOpenFile });

		(instance as unknown as { openFile(): void }).openFile();

		expect(onOpenFile).toHaveBeenCalledTimes(1);
	});
});

describe('vanilla options shared with the other bindings', () => {
	it('keys the recovery snapshot on filePath, ahead of the deprecated autosaveFilePath', async () => {
		const key = async (options: PptxViewerOptions): Promise<string> => {
			const instance = await mountLoaded(options);
			const resolved = (instance as unknown as { autosaveFilePath(): string }).autosaveFilePath();
			instance.destroy();
			viewer = undefined;
			return resolved;
		};

		expect(await key({ filePath: 'new.pptx', autosaveFilePath: 'old.pptx' })).toBe('new.pptx');
		expect(await key({ autosaveFilePath: 'old.pptx' })).toBe('old.pptx');
	});

	it('authors comments as authorName', async () => {
		const instance = await mountLoaded({ editable: true, authorName: 'Dana Host' });
		const { editor } = instance as unknown as {
			editor: { getEditActions(): { comments: { addComment(text: string): string | null } } };
		};

		editor.getEditActions().comments.addComment('Written by the host identity');

		expect(instance.getSlides()[0]?.comments?.[0]?.author).toBe('Dana Host');
	});

	it('broadcasts under authorName rather than the default presenter', async () => {
		const { buildBroadcastSessionConfig } = await import('./collab/broadcast-helpers');
		expect(
			buildBroadcastSessionConfig({ roomId: 'r', serverUrl: 'ws://x' }, 'Dana Host').userName,
		).toBe('Dana Host');
	});
});
