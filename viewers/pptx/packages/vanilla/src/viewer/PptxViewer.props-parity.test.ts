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

describe('vanilla setters for options the other bindings pass as live props', () => {
	it('shows and hides the toolbar without reloading the deck or losing the undo history', async () => {
		const instance = await mountLoaded({ editable: true });
		const root = document.body;
		expect(root.querySelector('.pptxv-ribbon')).not.toBeNull();
		instance.addSlide();
		const slides = instance.getSlideCount();
		expect(instance.canUndo()).toBe(true);

		instance.setShowToolbar(false);

		expect(root.querySelector('.pptxv-ribbon')).toBeNull();
		expect(instance.getSlideCount()).toBe(slides);
		expect(instance.canUndo()).toBe(true);

		instance.setShowToolbar(true);

		expect(root.querySelector('.pptxv-ribbon')).not.toBeNull();
		instance.undo();
		expect(instance.getSlideCount()).toBe(slides - 1);
	});

	it('starts from the showToolbar option and ignores a call that changes nothing', async () => {
		const instance = await mountLoaded({ showToolbar: false });
		expect(document.body.querySelector('.pptxv-ribbon')).toBeNull();
		const before = document.body.querySelector('.pptxv');

		instance.setShowToolbar(false);

		expect(document.body.querySelector('.pptxv-ribbon')).toBeNull();
		expect(document.body.querySelector('.pptxv')).toBe(before);
	});
});

describe('ribbon display options', () => {
	const shownPane = (): HTMLElement | undefined =>
		[...document.body.querySelectorAll<HTMLElement>('.pptxv-ribbon-tab-content')].find(
			(pane) => !pane.hidden,
		);
	const tab = (label: string): HTMLButtonElement =>
		[...document.body.querySelectorAll<HTMLButtonElement>('.pptxv-ribbon-tab')].find(
			(button) => button.textContent === label,
		)!;

	it('lists the three PowerPoint modes and collapses the ribbon to its tab row', async () => {
		const onRibbonDisplayChange = vi.fn();
		const instance = await mountLoaded({ editable: true, onRibbonDisplayChange });
		const ribbon = document.body.querySelector<HTMLElement>('.pptxv-ribbon')!;
		expect(instance.getRibbonDisplay()).toBe('always');
		expect(shownPane()).toBeDefined();

		ribbon.querySelector<HTMLButtonElement>('.pptxv-ribbon-display-btn')!.click();
		const items = [
			...ribbon.querySelectorAll<HTMLButtonElement>(
				'.pptxv-ribbon-display-menu [role="menuitemradio"]',
			),
		];
		expect(items.map((item) => item.textContent)).toEqual([
			'Full-screen mode',
			'Show tabs only',
			'Always show Ribbon',
		]);
		expect(items.map((item) => item.getAttribute('aria-checked'))).toEqual([
			'false',
			'false',
			'true',
		]);
		items[1]!.click();

		expect(onRibbonDisplayChange).toHaveBeenLastCalledWith('tabs');
		expect(instance.getRibbonDisplay()).toBe('tabs');
		expect(shownPane()).toBeUndefined();

		// A tab brings its commands up over the slide; Escape puts them away again.
		tab('Insert').click();
		expect(shownPane()?.hasAttribute('data-ribbon-overlay')).toBe(true);
		document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(shownPane()).toBeUndefined();

		// A double click on a tab docks the ribbon again, as Ctrl+F1 does.
		tab('Insert').dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		expect(onRibbonDisplayChange).toHaveBeenLastCalledWith('always');
		expect(shownPane()?.hasAttribute('data-ribbon-overlay')).toBe(false);
		document.body.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'F1', ctrlKey: true, bubbles: true }),
		);
		expect(instance.getRibbonDisplay()).toBe('tabs');
	});

	it('hides the ribbon behind a strip in full-screen mode and keeps the mode across a rebuild', async () => {
		const onRibbonDisplayChange = vi.fn();
		const instance = await mountLoaded({ ribbonDisplay: 'tabs', onRibbonDisplayChange });
		expect(instance.getRibbonDisplay()).toBe('tabs');

		instance.setRibbonDisplay('fullScreen');

		const root = (): HTMLElement => document.body.querySelector<HTMLElement>('.pptxv')!;
		expect(root().dataset.ribbonDisplay).toBe('fullScreen');
		const strip = root().querySelector<HTMLButtonElement>('.pptxv-ribbon-reveal')!;
		expect(strip.nextElementSibling?.classList.contains('pptxv-ribbon')).toBe(true);
		expect(onRibbonDisplayChange).not.toHaveBeenCalled();
		strip.click();
		expect(root().querySelector('.pptxv-ribbon')!.hasAttribute('data-ribbon-peek')).toBe(true);

		// A locale switch rebuilds the chrome; the mode is the viewer's, not the old ribbon's.
		instance.setLocale('en');
		expect(instance.getRibbonDisplay()).toBe('fullScreen');
		expect(document.body.querySelectorAll('.pptxv-ribbon-reveal')).toHaveLength(1);
		expect(root().dataset.ribbonDisplay).toBe('fullScreen');

		instance.setRibbonDisplay('always');
		expect(document.body.querySelector('.pptxv-ribbon-reveal')).toBeNull();
		expect(root().dataset.ribbonDisplay).toBe('always');

		// The ribbon that the rebuild replaced no longer listens: a press on the new ribbon's
		// menu must not be taken for a press outside it, which would close the menu first.
		const ribbon = root().querySelector<HTMLElement>('.pptxv-ribbon')!;
		ribbon.querySelector<HTMLButtonElement>('.pptxv-ribbon-display-btn')!.click();
		const item = ribbon.querySelector<HTMLButtonElement>(
			'.pptxv-ribbon-display-menu [data-ribbon-display="tabs"]',
		)!;
		item.dispatchEvent(new Event('pointerdown', { bubbles: true, composed: true }));
		expect(ribbon.querySelector<HTMLElement>('.pptxv-ribbon-display-menu')!.hidden).toBe(false);
		item.click();
		expect(instance.getRibbonDisplay()).toBe('tabs');
		expect(onRibbonDisplayChange).toHaveBeenLastCalledWith('tabs');
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
