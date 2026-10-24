import type { RibbonGalleryContext } from 'ooxml-ui/pptx';
import type { PptxHandler, PptxSlide } from 'pptx-viewer-core';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createGalleryActions } from '../../../editor/editor-gallery-actions';
import { createTranslator } from '../../../i18n';
import { createInitialViewerState, createStore } from '../../../state';
import { createRibbonGalleryHub } from './gallery-hub';
import { createRibbonGallery } from './ribbon-gallery';

/**
 * Design > Variants > Background Styles in the vanilla binding: a hovered tile is shown over the
 * slides without an undo step or a dirty flag, leaving restores the original, a pick commits one
 * undo step, and the footer command reaches the Design tab.
 */
beforeEach(() => vi.useFakeTimers());
afterEach(() => {
	vi.useRealTimers();
	document.body.replaceChildren();
});

const handler = {
	applyBackgroundStyleToSlides: (slides: PptxSlide[], templates: unknown, index: number) => ({
		slides: slides.map((slide) => ({ ...slide, backgroundColor: `#00000${index}` })),
		templateElementsBySlideId: templates,
	}),
	commitMasterBackgroundStyle: vi.fn(async () => true),
} as unknown as PptxHandler;

function setup(onCommand = vi.fn()) {
	const slide = {
		id: 's1',
		rId: 'r1',
		slideNumber: 1,
		elements: [],
		backgroundColor: '#FFFFFF',
	} as unknown as PptxSlide;
	const store = createStore({
		...createInitialViewerState(),
		editable: true,
		slides: [slide],
	});
	const ops = { pushHistory: vi.fn(), commitChange: vi.fn() };
	const actions = createGalleryActions({
		store,
		ops,
		getHandler: () => handler,
	});
	const hub = createRibbonGalleryHub((result) => actions.applyRibbonGalleryResult(result), {
		preview: (result) => actions.previewRibbonGalleryResult(result),
		end: () => actions.endRibbonGalleryPreview(),
	});
	const gallery = createRibbonGallery(
		document,
		createTranslator(),
		{ gallery: 'themeBackgrounds', control: 'design.variants.backgroundStyles', mode: 'dropdown' },
		hub,
		{ onCommand },
	);
	document.body.append(gallery.el);
	const context: RibbonGalleryContext = { element: null, theme: { name: 'Office' } };
	hub.sync(context, true);
	gallery.trigger.click();
	return { store, ops, gallery, slide };
}

describe('Design > Background Styles (vanilla)', () => {
	it('shows a hovered style over the slides and restores the original on leave', () => {
		const { store, ops, gallery, slide } = setup();
		const tile = gallery.popup.querySelector('[data-gallery-item="3"]')!;
		tile.dispatchEvent(new PointerEvent('pointerenter', { pointerType: 'mouse' }));
		vi.advanceTimersByTime(150);
		expect(store.get().slides[0]!.backgroundColor).toBe('#000003');
		expect(store.get().deckStyleHeld).toBe(true);
		tile.dispatchEvent(new PointerEvent('pointerleave', { pointerType: 'mouse' }));
		vi.advanceTimersByTime(100);
		expect(store.get().slides[0]).toBe(slide);
		expect(store.get().deckStyleHeld).toBe(false);
		expect(ops.pushHistory).not.toHaveBeenCalled();
		expect(ops.commitChange).not.toHaveBeenCalled();
		expect(store.get().dirty).toBe(false);
	});

	it('commits a pick as one undo step from the real slides', async () => {
		const { store, ops, gallery } = setup();
		const tile = gallery.popup.querySelector<HTMLButtonElement>('[data-gallery-item="4"]')!;
		tile.dispatchEvent(new PointerEvent('pointerenter', { pointerType: 'mouse' }));
		vi.advanceTimersByTime(150);
		tile.click();
		await vi.advanceTimersByTimeAsync(0);
		expect(ops.pushHistory).toHaveBeenCalledOnce();
		expect(store.get().slides[0]!.backgroundColor).toBe('#000004');
		expect(ops.commitChange).toHaveBeenCalledOnce();
		expect(store.get().deckStyleHeld).toBe(false);
	});

	it('runs the Format Background command from the footer', () => {
		const onCommand = vi.fn();
		const { gallery } = setup(onCommand);
		gallery.popup
			.querySelector<HTMLButtonElement>('[data-gallery-footer="formatBackground"]')!
			.click();
		expect(onCommand).toHaveBeenCalledWith('formatBackground');
	});
});
