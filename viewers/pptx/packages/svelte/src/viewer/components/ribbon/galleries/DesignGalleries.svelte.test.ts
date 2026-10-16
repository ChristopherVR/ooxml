import { FIXED_TAB_GALLERIES } from 'ooxml-ui/pptx';
import type { RibbonGalleryPlacement } from 'ooxml-ui/pptx';
import type { PptxHandler } from 'pptx-viewer-core';
import { THEME_PRESETS } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { EditorState } from '../../../editor/editor-state.svelte';
import { createRibbonGalleryHost, ribbonGalleryHostContext } from './ribbon-gallery-host';
import RibbonGallery from './RibbonGallery.svelte';

/**
 * Design galleries: a hovered Background Styles tile is shown over the slides without an undo
 * step or a dirty flag, leaving puts the original back, and a pick commits as one undo step.
 */
let cleanup: (() => void) | undefined;

beforeEach(() => vi.useFakeTimers());
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
	vi.useRealTimers();
});

const BACKGROUNDS = FIXED_TAB_GALLERIES.find(
	(placement) => placement.gallery === 'themeBackgrounds',
) as RibbonGalleryPlacement;

function fakeHandler(): PptxHandler {
	return {
		resolveStyleMatrixReferences: () => ({ shapeStyle: {} }),
		getThemeVariants: () => [],
		applyBackgroundStyleToSlides: (slides: never[], templates: unknown, index: number) => ({
			slides: slides.map((slide) => ({ ...(slide as object), backgroundColor: `#00000${index}` })),
			templateElementsBySlideId: templates,
		}),
		commitMasterBackgroundStyle: vi.fn(async () => true),
	} as unknown as PptxHandler;
}

function makeEditor(): EditorState {
	const handler = fakeHandler();
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => handler });
	editor.editable = true;
	editor.theme = {
		name: 'Office',
		colorScheme: THEME_PRESETS[0].colorScheme,
		fontScheme: THEME_PRESETS[0].fontScheme,
	};
	editor.setSlides([
		{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [], backgroundColor: '#FFFFFF' },
	]);
	return editor;
}

function mountGallery(editor: EditorState): HTMLElement {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(RibbonGallery, {
		target,
		props: { placement: BACKGROUNDS },
		context: ribbonGalleryHostContext(createRibbonGalleryHost(editor)),
	});
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	target.querySelector<HTMLButtonElement>('[data-ribbon-gallery]')!.click();
	flushSync();
	return target;
}

const hover = (tile: Element): void => {
	tile.dispatchEvent(new PointerEvent('pointerenter', { pointerType: 'mouse' }));
	vi.advanceTimersByTime(150);
	flushSync();
};

describe('Design > Background Styles gallery', () => {
	it('shows a hovered style over the slides and puts the original back on leave', () => {
		const editor = makeEditor();
		const original = editor.slides;
		const target = mountGallery(editor);
		const tile = target.querySelector('[data-gallery-item="3"]')!;
		hover(tile);
		expect(editor.slides[0]!.backgroundColor).toBe('#000003');
		expect(editor.deckStyleHeld).toBe(true);
		expect(editor.dirty).toBe(false);
		expect(editor.committedRenderedSlides[0]!.backgroundColor).toBe('#FFFFFF');
		tile.dispatchEvent(new PointerEvent('pointerleave', { pointerType: 'mouse' }));
		vi.advanceTimersByTime(100);
		flushSync();
		expect(editor.slides).toBe(original);
		expect(editor.deckStyleHeld).toBe(false);
		expect(editor.dirty).toBe(false);
	});

	it('commits a pick as one undoable step from the real slides', async () => {
		const editor = makeEditor();
		const target = mountGallery(editor);
		hover(target.querySelector('[data-gallery-item="4"]')!);
		target.querySelector<HTMLButtonElement>('[data-gallery-item="4"]')!.click();
		await vi.advanceTimersByTimeAsync(0);
		flushSync();
		expect(editor.slides[0]!.backgroundColor).toBe('#000004');
		expect(editor.dirty).toBe(true);
		expect(editor.deckStyleHeld).toBe(false);
		editor.undo();
		expect(editor.slides[0]!.backgroundColor).toBe('#FFFFFF');
	});

	it('lists the tile menu and the Format Background footer', () => {
		const target = mountGallery(makeEditor());
		expect(target.querySelectorAll('[data-gallery-item]')).toHaveLength(12);
		expect(target.querySelector('[data-gallery-footer="formatBackground"]')).not.toBeNull();
	});
});
