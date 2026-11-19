import { DEFERRED_FILL_SETTLE_MS, editorThumbnailStep } from 'ooxml-ui/pptx';
/**
 * Thumbnail-rail tests for the display contract: the rail must never carry an
 * inline `display` style, because presentation mode and the mobile layout hide
 * it with stylesheet rules (`.pptxv-presenting .pptxv-thumbs` / the mobile
 * media query) that inline styles would override, leaking thumbnail content
 * into the presented slide show.
 */
import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createThumbnailRail } from './thumbnails';

const CANVAS = { width: 960, height: 540 };

function slideRenderer(slide: PptxSlide): HTMLElement {
	const el = document.createElement('div');
	el.dataset.slideId = slide.id;
	return el;
}

function makeSlides(count: number): PptxSlide[] {
	return Array.from({ length: count }, (_, index) => ({
		id: `slide-${index}`,
		rId: `rId-${index}`,
		slideNumber: index + 1,
		elements: [],
	})) as PptxSlide[];
}

describe('thumbnail rail display contract', () => {
	it('moves the current-page cue and delegates row spacing to the shared window', () => {
		const onSelect = vi.fn();
		const rail = createThumbnailRail(document, createTranslator(), onSelect);
		rail.render(makeSlides(3), CANVAS, slideRenderer);
		const window = rail.el.querySelector<HTMLElement>('[data-pptx-chrome="slide-window"]')!;
		expect(window.style.gap).toBe('');
		const rows = window.querySelectorAll<HTMLButtonElement>('[data-pptx-chrome="slide-row"]');
		rail.setActive(1);
		expect(rows[0].hasAttribute('aria-current')).toBeFalsy();
		// Same token as the React, Vue, Svelte and Angular rails.
		expect(rows[1].getAttribute('aria-current')).toBe('true');
		expect(rows[1].querySelector('[data-pptx-chrome="slide-number"]')?.textContent).toBe('2');
		rows[2].click();
		expect(onSelect).toHaveBeenCalledExactlyOnceWith(2);
	});

	it('preserves slide aspect ratio inside the shared desktop rail', () => {
		const rail = createThumbnailRail(document, createTranslator(), vi.fn());
		rail.render(makeSlides(3), CANVAS, slideRenderer);
		const frame = rail.el.querySelector<HTMLElement>('[data-pptx-chrome="slide-frame"]');
		expect(rail.el.style.width).toBe('180px');
		expect(frame?.style.width).toBe('132px');
		expect(frame?.style.height).toBe('74.25px');
	});

	it('never sets an inline display style (small deck)', () => {
		const rail = createThumbnailRail(document, createTranslator(), vi.fn());
		rail.render(makeSlides(3), CANVAS, slideRenderer);
		expect(rail.el.style.display).toBe('');
		expect(rail.el.classList.contains('pptxv-thumbs-virtualized')).toBeFalsy();
	});

	it('switches to the virtualized class (not an inline style) for large decks', () => {
		const rail = createThumbnailRail(document, createTranslator(), vi.fn());
		rail.render(makeSlides(100), CANVAS, slideRenderer);
		expect(rail.el.style.display).toBe('');
		expect(rail.el.classList.contains('pptxv-thumbs-virtualized')).toBeTruthy();
	});

	it('drops the virtualized class again when the deck shrinks', () => {
		const rail = createThumbnailRail(document, createTranslator(), vi.fn());
		rail.render(makeSlides(100), CANVAS, slideRenderer);
		rail.render(makeSlides(3), CANVAS, slideRenderer);
		expect(rail.el.classList.contains('pptxv-thumbs-virtualized')).toBeFalsy();
	});

	it('drops the virtualized class in master view', () => {
		const rail = createThumbnailRail(document, createTranslator(), vi.fn());
		rail.render(makeSlides(100), CANVAS, slideRenderer);
		rail.renderMasters(
			[{ path: 'master-1', name: 'Corporate', elements: [], layouts: [] }],
			CANVAS,
			slideRenderer,
			vi.fn(),
			{ masterIndex: 0, layoutIndex: null },
		);
		expect(rail.el.style.display).toBe('');
		expect(rail.el.classList.contains('pptxv-thumbs-virtualized')).toBeFalsy();
	});

	it('setVisible toggles the hidden attribute', () => {
		const rail = createThumbnailRail(document, createTranslator(), vi.fn());
		rail.setVisible(false);
		expect(rail.el.hidden).toBeTruthy();
		rail.setVisible(true);
		expect(rail.el.hidden).toBeFalsy();
	});
});

describe('thumbnail rail windowing', () => {
	const STEP = editorThumbnailStep(CANVAS.width, CANVAS.height);
	const nextFrame = () =>
		new Promise<void>((resolve) => {
			requestAnimationFrame(() => resolve());
		});
	const rowsOf = (rail: ReturnType<typeof createThumbnailRail>) => [
		...rail.el.querySelectorAll<HTMLButtonElement>('[data-pptx-chrome="slide-row"]'),
	];
	const mount = () => {
		const renderStage = vi.fn(slideRenderer);
		const rail = createThumbnailRail(document, createTranslator(), vi.fn());
		document.body.appendChild(rail.el);
		rail.render(makeSlides(200), CANVAS, renderStage);
		const list = rail.el.querySelector<HTMLElement>('[data-pptx-chrome="slide-list"]')!;
		return { rail, list, renderStage };
	};

	it('builds only the rows entering the window, once per frame of scrolling', async () => {
		const { rail, list, renderStage } = mount();
		const before = rowsOf(rail);
		renderStage.mockClear();
		for (const rows of [1, 2, 3]) {
			list.scrollTop = rows * STEP;
			list.dispatchEvent(new Event('scroll'));
		}
		expect(renderStage).not.toHaveBeenCalled();
		await nextFrame();
		// Rows still in the window are the same nodes; three row shells were added,
		// and their thumbnails wait for scrolling to pause.
		const after = rowsOf(rail);
		expect(after).toHaveLength(before.length + 3);
		expect(after.slice(0, before.length)).toStrictEqual(before);
		expect(renderStage).not.toHaveBeenCalled();
		await new Promise((resolve) => {
			setTimeout(resolve, DEFERRED_FILL_SETTLE_MS + 50);
		});
		await nextFrame();
		expect(renderStage).toHaveBeenCalledTimes(3);
		expect(after.at(-1)?.querySelector('.pptxv-thumb-pending')).toBeNull();
		rail.destroy();
		rail.el.remove();
	});

	it('moves the current slide without rebuilding rows while the window holds', () => {
		const { rail, renderStage } = mount();
		const before = rowsOf(rail);
		renderStage.mockClear();
		rail.setActive(2);
		expect(renderStage).not.toHaveBeenCalled();
		expect(rowsOf(rail)).toStrictEqual(before);
		expect(before[2]?.getAttribute('aria-current')).toBe('true');
		expect(before[0]?.hasAttribute('aria-current')).toBeFalsy();
		rail.destroy();
		rail.el.remove();
	});

	it('builds only the entering row when the current slide moves the window', () => {
		const { rail, list, renderStage } = mount();
		const visible = Math.floor(600 / STEP);
		renderStage.mockClear();
		rail.setActive(visible + 1);
		expect(list.scrollTop).toBeGreaterThan(0);
		expect(renderStage.mock.calls.length).toBeLessThanOrEqual(2);
		const current = rail.el.querySelector<HTMLElement>('[aria-current="true"]');
		expect(current?.dataset.slideIndex).toBe(String(visible + 1));
		rail.destroy();
		rail.el.remove();
	});

	it('keeps keyboard focus on the current slide as paging moves the window', () => {
		const { rail } = mount();
		rowsOf(rail)[0]!.focus();
		for (let index = 1; index <= 30; index += 1) {
			rail.setActive(index);
		}
		const focused = document.activeElement as HTMLElement | null;
		expect(focused?.getAttribute('aria-current')).toBe('true');
		expect(focused?.querySelector('[data-pptx-chrome="slide-number"]')?.textContent).toBe('31');
		rail.destroy();
		rail.el.remove();
	});
});
