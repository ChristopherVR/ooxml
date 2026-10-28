/**
 * `createSlidesPane`: the slide thumbnail rail as a host-driven building
 * block. Pins the DOM markers the cross-binding rail specs read, that the
 * thumbnails draw what the canvas draws (deck start number included), and
 * that every rail command reaches the host's callbacks.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { createSlidesPaneBulkSlideOps } from 'ooxml-ui/pptx';
import type { FieldSubstitutionContext } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSection, PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import * as root from '../../index';
import { loadPresentation, revokeBlobUrls } from '../load';
import type { LoadedPresentation } from '../load';
import { createDefaultRegistry } from '../render';
import { buildRenderFieldContext } from '../render-field-context';
import { createInitialViewerState, createStore } from '../state';
import type { ViewerState } from '../state';
import { menuItem } from './context-menu.test-support';
import { createSlidesPane } from './slides-pane';
import type { SlidesPaneOptions } from './slides-pane';

const CANVAS = { width: 960, height: 540 };

function makeSlides(count: number): PptxSlide[] {
	return Array.from({ length: count }, (_, index) => ({
		id: `s${index}`,
		rId: `rId${index}`,
		slideNumber: index + 1,
		elements: [],
	})) as PptxSlide[];
}

function textBox(id: string, text: string): PptxElement {
	return { id, type: 'text', x: 10, y: 10, width: 200, height: 40, text } as PptxElement;
}

function mount(overrides: Partial<SlidesPaneOptions> = {}) {
	const container = document.createElement('div');
	container.className = 'pptxv';
	document.body.appendChild(container);
	const onSelectSlide = vi.fn();
	const pane = createSlidesPane(container, {
		slides: makeSlides(4),
		activeSlideIndex: 0,
		canvasSize: CANVAS,
		onSelectSlide,
		...overrides,
	});
	return { container, pane, onSelectSlide };
}

const rows = (scope: ParentNode): HTMLButtonElement[] =>
	Array.from(scope.querySelectorAll<HTMLButtonElement>('[data-pptx-chrome="slide-row"]'));

function rightClick(target: HTMLElement, x = 40, y = 120): void {
	target.dispatchEvent(
		new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: x, clientY: y }),
	);
}

const deck = (): ArrayBuffer => {
	const bytes = readFileSync(
		resolve(process.cwd(), '../../../../e2e/pptx/fixtures/first-slide-number.pptx'),
	);
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};

let loaded: LoadedPresentation | undefined;
afterEach(() => {
	if (loaded) {
		loaded.handler.dispose();
		revokeBlobUrls(loaded.blobUrls);
		loaded = undefined;
	}
	document.body.replaceChildren();
});

describe('createSlidesPane', () => {
	it('is exported from the package root with the bulk slide helper', () => {
		expect(root.createSlidesPane).toBe(createSlidesPane);
		expect(root.createSlidesPaneBulkSlideOps).toBe(createSlidesPaneBulkSlideOps);
	});

	it('emits the rail markers shared with the other bindings', () => {
		const { container, pane } = mount({
			activeSlideIndex: 1,
			canEdit: true,
			onAddSlide: vi.fn(),
		});
		expect(pane.el.parentElement).toBe(container);
		expect(pane.el.dataset.pptxChrome).toBe('slides');
		const list = rows(container);
		expect(list).toHaveLength(4);
		expect(list[1].getAttribute('aria-current')).toBe('true');
		expect(list.filter((row) => row.hasAttribute('aria-current'))).toHaveLength(1);
		expect(list[0].getAttribute('aria-label')).toBe('Go to slide 1');
		expect(list[0].querySelector('[data-pptx-chrome="slide-frame"]')).not.toBeNull();
		const footer = container.querySelector<HTMLElement>('[data-pptx-chrome="slide-footer"]');
		expect(footer?.hidden).toBeFalsy();
		expect(footer?.querySelector('button')).not.toBeNull();
	});

	it('numbers thumbnails from the deck start number (firstSlideNum="45")', async () => {
		loaded = await loadPresentation(deck());
		expect(loaded.firstSlideNumber).toBe(45);
		const { container } = mount({
			slides: loaded.slides,
			canvasSize: loaded.canvasSize,
			mediaDataUrls: loaded.mediaDataUrls,
			headerFooter: loaded.headerFooter,
			customProperties: loaded.customProperties,
			firstSlideNumber: loaded.firstSlideNumber,
		});
		const frames = Array.from(
			container.querySelectorAll<HTMLElement>('[data-pptx-chrome="slide-frame"]'),
		).map((frame) => frame.textContent ?? '');
		expect(frames).toHaveLength(3);
		expect(frames[0]).toContain('Slide 45');
		expect(frames[1]).toContain('Slide 46');
		expect(frames[2]).toContain('Slide 47');
		expect(frames.join(' ')).not.toContain('Slide 1');
	});

	it('hands each thumbnail the field context the canvas uses for that slide', async () => {
		loaded = await loadPresentation(deck());
		const registry = createDefaultRegistry();
		const seen = new Map<string, FieldSubstitutionContext | undefined>();
		const base = registry.resolve('text');
		registry.register('text', (element, zIndex, context) => {
			seen.set(element.id, context.fieldContext);
			return base(element, zIndex, context);
		});
		const result = loaded;
		mount({
			slides: result.slides,
			canvasSize: result.canvasSize,
			headerFooter: result.headerFooter,
			customProperties: result.customProperties,
			firstSlideNumber: result.firstSlideNumber,
			registry,
		});
		const canvasState = {
			headerFooter: result.headerFooter,
			customProperties: result.customProperties,
			firstSlideNumber: result.firstSlideNumber,
		};
		let checked = 0;
		for (const slide of result.slides) {
			const expected = buildRenderFieldContext(canvasState, slide);
			for (const element of slide.elements) {
				if (seen.has(element.id)) {
					expect(seen.get(element.id)).toStrictEqual(expected);
					checked += 1;
				}
			}
		}
		expect(checked).toBeGreaterThanOrEqual(3);
	});

	it('draws template elements behind the slide like the canvas', () => {
		const slides = makeSlides(2);
		slides[0] = { ...slides[0], elements: [textBox('own', 'Own text')] };
		const { container } = mount({
			slides,
			templateElementsBySlideId: { s0: [textBox('layout', 'Layout text')] },
		});
		const frame = rows(container)[0].querySelector('[data-pptx-chrome="slide-frame"]');
		const ids = Array.from(frame?.querySelectorAll<HTMLElement>('[data-element-id]') ?? []).map(
			(node) => node.dataset.elementId,
		);
		expect(ids).toStrictEqual(['layout', 'own']);
	});

	it('selecting moves the active slide without editing the deck', () => {
		const store = createStore<ViewerState>({
			...createInitialViewerState(),
			slides: makeSlides(3),
		});
		const before = store.get().slides;
		const container = document.createElement('div');
		document.body.appendChild(container);
		const pane = createSlidesPane(container, {
			slides: store.get().slides,
			activeSlideIndex: store.get().currentSlide,
			canvasSize: CANVAS,
			canEdit: true,
			onSelectSlide: (index) => store.set({ currentSlide: index }),
		});
		store.subscribe((state) => pane.update({ activeSlideIndex: state.currentSlide }));
		const frame = rows(container)[2].querySelector('[data-pptx-chrome="slide-frame"]');

		rows(container)[2].click();

		expect(store.get().currentSlide).toBe(2);
		expect(store.get().dirty).toBeFalsy();
		expect(store.get().slides).toBe(before);
		expect(rows(container)[2].getAttribute('aria-current')).toBe('true');
		expect(rows(container)[0].hasAttribute('aria-current')).toBeFalsy();
		// Only the current marker moved: the thumbnails were not redrawn.
		expect(rows(container)[2].querySelector('[data-pptx-chrome="slide-frame"]')).toBe(frame);
	});

	it('the Layout entry makes the slide active, then asks the host for its gallery', () => {
		const calls: string[] = [];
		const onOpenLayoutGallery = vi.fn(() => {
			calls.push('gallery');
		});
		const { container, onSelectSlide } = mount({
			canEdit: true,
			onOpenLayoutGallery,
		});
		onSelectSlide.mockImplementation(() => {
			calls.push('select');
		});

		rightClick(rows(container)[1]);
		menuItem('layout')?.click();

		expect(onSelectSlide).toHaveBeenCalledExactlyOnceWith(1);
		expect(onOpenLayoutGallery).toHaveBeenCalledExactlyOnceWith({ slideIndex: 1, x: 40, y: 120 });
		expect(calls).toStrictEqual(['select', 'gallery']);
		expect(menuItem('layout')).toBeNull();
	});

	it('the Layout entry still selects the slide without a host gallery', () => {
		const { container, onSelectSlide } = mount({ canEdit: true });
		rightClick(rows(container)[3]);
		menuItem('layout')?.click();
		expect(onSelectSlide).toHaveBeenCalledExactlyOnceWith(3);
	});

	it('bulk duplicate copies every selected slide right after its source', () => {
		let slides = makeSlides(4);
		const ops = createSlidesPaneBulkSlideOps({
			duplicateSlide: (index) => {
				const copy = { ...slides[index], id: `${slides[index].id}-copy` };
				slides = [...slides.slice(0, index + 1), copy, ...slides.slice(index + 1)];
			},
			deleteSlide: vi.fn(),
			toggleSlideHidden: vi.fn(),
		});
		const onDuplicateSlides = vi.fn(ops.duplicateSlides);
		const { container } = mount({ slides, canEdit: true, onDuplicateSlides });

		rows(container)[0].click();
		rows(container)[2].dispatchEvent(new MouseEvent('click', { bubbles: true, ctrlKey: true }));
		rightClick(rows(container)[2]);
		menuItem('duplicate')?.click();

		expect(onDuplicateSlides).toHaveBeenCalledOnce();
		expect([...onDuplicateSlides.mock.calls[0][0]].sort()).toStrictEqual([0, 2]);
		expect(slides.map((slide) => slide.id)).toStrictEqual([
			's0',
			's0-copy',
			's1',
			's2',
			's2-copy',
			's3',
		]);
	});

	it('bulk delete and hide act on every selected slide', () => {
		const onDeleteSlides = vi.fn();
		const onHideSlides = vi.fn();
		const { container } = mount({ canEdit: true, onDeleteSlides, onHideSlides });

		rows(container)[1].click();
		rows(container)[3].dispatchEvent(new MouseEvent('click', { bubbles: true, ctrlKey: true }));
		rightClick(rows(container)[3]);
		menuItem('delete')?.click();
		expect([...onDeleteSlides.mock.calls[0][0]].sort()).toStrictEqual([1, 3]);

		rightClick(rows(container)[1]);
		menuItem('hide')?.click();
		expect([...onHideSlides.mock.calls[0][0]].sort()).toStrictEqual([1, 3]);
	});

	it('the footer and Enter go through the host add operations', () => {
		const onAddSlide = vi.fn();
		const onAddSlideAfter = vi.fn();
		const { container } = mount({
			activeSlideIndex: 2,
			canEdit: true,
			onAddSlide,
			onAddSlideAfter,
		});
		container.querySelector<HTMLButtonElement>('[data-pptx-chrome="slide-footer"] button')!.click();
		expect(onAddSlide).toHaveBeenCalledOnce();
		rows(container)[2].dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		expect(onAddSlideAfter).toHaveBeenCalledExactlyOnceWith(2);
	});

	it('read-only shows no footer, no menus and no Enter insert', () => {
		const onAddSlide = vi.fn();
		const onAddSlideAfter = vi.fn();
		const { container, pane, onSelectSlide } = mount({ onAddSlide, onAddSlideAfter });
		const footer = container.querySelector<HTMLElement>('[data-pptx-chrome="slide-footer"]');
		expect(footer === null || footer.hidden).toBeTruthy();
		rightClick(rows(container)[1]);
		expect(menuItem('layout')).toBeNull();
		rows(container)[1].dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		expect(onAddSlideAfter).not.toHaveBeenCalled();
		rows(container)[1].click();
		expect(onSelectSlide).toHaveBeenCalledExactlyOnceWith(1);

		// Turning editing on later brings the footer and the menus back.
		pane.update({ canEdit: true });
		expect(container.querySelectorAll('[data-pptx-chrome="slides"]')).toHaveLength(1);
		const shown = container.querySelector<HTMLElement>('[data-pptx-chrome="slide-footer"]');
		expect(shown).not.toBeNull();
		expect(shown?.hidden).toBeFalsy();
		rightClick(rows(container)[1]);
		expect(menuItem('layout')).not.toBeNull();
	});

	it('maps the section headers onto the host section callbacks', () => {
		const slides = makeSlides(3).map((slide, index) => ({
			...slide,
			sectionId: index === 0 ? 'sec1' : 'sec2',
		}));
		const sections: PptxSection[] = [
			{ id: 'sec1', name: 'Intro', slideIds: ['s0'] },
			{ id: 'sec2', name: 'Body', slideIds: ['s1', 's2'] },
		];
		const onToggleSection = vi.fn();
		const onAddSection = vi.fn();
		const { container } = mount({
			slides,
			sections,
			canEdit: true,
			onToggleSection,
			onAddSection,
		});
		const headers = container.querySelectorAll<HTMLElement>('[data-pptx-chrome="section-header"]');
		expect(headers).toHaveLength(2);
		expect(rows(container)).toHaveLength(3);
		headers[0].querySelector<HTMLButtonElement>('button')!.click();
		expect(onToggleSection).toHaveBeenCalledExactlyOnceWith('sec1');

		rightClick(rows(container)[2]);
		menuItem('add-section')?.click();
		expect(onAddSection).toHaveBeenCalledExactlyOnceWith('Untitled Section', 2);
	});

	it('redraws on a slide change and is removed by destroy', () => {
		const { container, pane } = mount();
		pane.update({ slides: makeSlides(6) });
		expect(rows(container)).toHaveLength(6);
		pane.destroy();
		expect(container.querySelector('[data-pptx-chrome="slides"]')).toBeNull();
		pane.update({ slides: makeSlides(2) });
		expect(container.childElementCount).toBe(0);
	});
});
