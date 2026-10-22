/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The fitted slide at fractional viewport sizes, under ancestor transforms and CSS zoom, and
 * after padding-only changes.
 *
 * - With no fit padding (`?fitPadding=0`, see `demos/pptx/shared/demo-root-options.ts`), a
 *   viewport laid out at a fractional size must hold the whole slide. Every binding used to fit
 *   to the whole-pixel `clientWidth` / `clientHeight`, which rounds up by up to half a pixel, so
 *   the slide overflowed the viewport by that much.
 * - Under an ancestor `transform` or CSS `zoom` the slide is laid out and painted through the
 *   scale, and its own layout rounding moves its edges by a fraction of a pixel, so an exact fit
 *   can still overflow (and, under a CSS zoom, make the viewport scroll by a pixel). There the
 *   fit reserves a pixel; all four edges must stay inside.
 * - A content-box viewport's client size also changes with its padding alone, which a default
 *   (content-box) `ResizeObserver` does not see; the fit must follow it.
 *
 * The viewport is pinned with injected CSS, so the same steps run unchanged against all five
 * demos.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { loadDeck, SAMPLE_DECK } from './support/deck';

test.use({ viewport: { width: 1600, height: 1000 } });

/** A layout unit (1/64px), the finest step a browser lays boxes out in. */
const LAYOUT_UNIT = 1 / 64;

interface Size {
	width: number;
	height: number;
}

/** Pin the viewport to `size` with injected CSS, plus any `extraCss`. */
async function pin(page: Page, size: Size, extraCss = ''): Promise<void> {
	await page.addInitScript(
		({ pinned, extra }) => {
			const style = document.createElement('style');
			style.id = 'pinned-viewport';
			style.textContent = `[data-pptx-viewport]{flex:none!important;width:${pinned.width}px!important;height:${pinned.height}px!important;min-width:0!important;min-height:0!important;max-width:none!important;max-height:none!important}${extra}`;
			const add = () => document.head.append(style);
			if (document.head) {
				add();
			} else {
				document.addEventListener('DOMContentLoaded', add);
			}
		},
		{ pinned: size, extra: extraCss },
	);
}

/** Let any late measurement land. */
async function settle(page: Page): Promise<void> {
	await page.evaluate(
		() =>
			new Promise((resolve) =>
				requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(resolve, 50))),
			),
	);
}

/**
 * The settled geometry, in the viewport's own CSS pixels. `transformScale` is the scale of an
 * ancestor `transform`; a CSS zoom is read from `currentCSSZoom`.
 */
async function settledGeometry(page: Page, transformScale = 1) {
	await expect(
		page.locator('[data-pptx-viewport] [data-pptx-element="true"]').first(),
	).toBeVisible();
	await settle(page);
	return page.evaluate((transform) => {
		const viewport = document.querySelector<HTMLElement>('[data-pptx-viewport]')!;
		const slide = viewport.querySelector('[aria-roledescription="slide"]')!;
		const zoom = (viewport as HTMLElement & { currentCSSZoom?: number }).currentCSSZoom ?? 1;
		const scale = transform * zoom;
		const box = viewport.getBoundingClientRect();
		const rect = slide.getBoundingClientRect();
		const style = getComputedStyle(viewport);
		const px = (value: string) => Number.parseFloat(value) || 0;
		// The content box, in viewport coordinates.
		const inset = {
			left: px(style.borderLeftWidth) + px(style.paddingLeft),
			right: px(style.borderRightWidth) + px(style.paddingRight),
			top: px(style.borderTopWidth) + px(style.paddingTop),
			bottom: px(style.borderBottomWidth) + px(style.paddingBottom),
		};
		const content = {
			left: box.left + inset.left * scale,
			right: box.right - inset.right * scale,
			top: box.top + inset.top * scale,
			bottom: box.bottom - inset.bottom * scale,
		};
		// How far each slide edge is past the content box (positive is outside).
		const edges = {
			left: (content.left - rect.left) / scale,
			right: (rect.right - content.right) / scale,
			top: (content.top - rect.top) / scale,
			bottom: (rect.bottom - content.bottom) / scale,
		};
		viewport.scrollLeft = 1e6;
		viewport.scrollTop = 1e6;
		const reachable = { left: viewport.scrollLeft, top: viewport.scrollTop };
		viewport.scrollLeft = 0;
		viewport.scrollTop = 0;
		return {
			content: {
				width: (content.right - content.left) / scale,
				height: (content.bottom - content.top) / scale,
			},
			slide: { width: rect.width / scale, height: rect.height / scale },
			maxEdge: Math.max(edges.left, edges.right, edges.top, edges.bottom),
			scrollRange: {
				width: viewport.scrollWidth - viewport.clientWidth,
				height: viewport.scrollHeight - viewport.clientHeight,
			},
			reachable,
		};
	}, transformScale);
}

type Geometry = Awaited<ReturnType<typeof settledGeometry>>;

/** The slide stays inside the content box on all four edges, and nothing can scroll. */
function expectContained(geometry: Geometry): void {
	expect(geometry.maxEdge).toBeLessThanOrEqual(LAYOUT_UNIT);
	expect(geometry.scrollRange).toStrictEqual({ width: 0, height: 0 });
	expect(geometry.reachable).toStrictEqual({ left: 0, top: 0 });
}

/** How much of the limiting axis the slide leaves free. */
function shortfall(geometry: Geometry): number {
	return Math.min(
		geometry.content.width - geometry.slide.width,
		geometry.content.height - geometry.slide.height,
	);
}

test.describe('viewport fit', () => {
	for (const [width, height] of [
		[801.5, 700],
		[1000.6, 800],
		[1400, 451.5],
	] as const) {
		test(`a ${width} x ${height} viewport holds the whole fitted slide`, async ({ page }) => {
			await pin(page, { width, height });
			await loadDeck(page, SAMPLE_DECK, '/?fitPadding=0');
			const geometry = await settledGeometry(page);
			expect(geometry.content.width).toBeCloseTo(width, 1);
			expect(geometry.content.height).toBeCloseTo(height, 1);
			expectContained(geometry);
			// Fitted exactly to the limiting axis, not shrunk by a whole pixel.
			expect(shortfall(geometry)).toBeLessThan(0.1);
		});
	}

	/** Scale the viewport's parent; the layout, and so the slide's own CSS pixels, are unchanged. */
	const SCALE = 1.0005;
	const SCALED_PARENT = `:has(> [data-pptx-viewport]){transform:scale(${SCALE})!important;transform-origin:0 0!important}`;

	test('an ancestor transform: scale(1.0005) does not push the slide past the viewport', async ({
		page,
	}) => {
		await pin(page, { width: 1000.4, height: 800 }, SCALED_PARENT);
		await loadDeck(page, SAMPLE_DECK, '/?fitPadding=0');
		const geometry = await settledGeometry(page, SCALE);
		expectContained(geometry);
		// A pixel is reserved under a transform, no more.
		expect(shortfall(geometry)).toBeLessThanOrEqual(1.5);
	});

	for (const boxSizing of ['border-box', 'content-box'] as const) {
		test(`a padded, bordered ${boxSizing} viewport under scale(1.0005) holds the slide in its content box`, async ({
			page,
		}) => {
			// The CSS padding equals the fit padding, so the slide fills the content box.
			const padding = 12.3;
			await pin(
				page,
				{ width: 1000.4, height: 700 },
				`${SCALED_PARENT}[data-pptx-viewport]{box-sizing:${boxSizing}!important;border:1.5px solid #c00!important;padding:${padding}px!important}`,
			);
			await loadDeck(page, SAMPLE_DECK, `/?fitPadding=${padding}`);
			const geometry = await settledGeometry(page, SCALE);
			expectContained(geometry);
			expect(shortfall(geometry)).toBeLessThanOrEqual(1.5);
		});
	}

	for (const zoom of [1.1, 0.5]) {
		test(`an ancestor zoom: ${zoom} keeps all four slide edges inside the viewport`, async ({
			page,
		}) => {
			await pin(
				page,
				{ width: 801.5, height: 700 },
				`:has(> [data-pptx-viewport]){zoom:${zoom}!important}[data-pptx-viewport]{box-sizing:border-box!important;border:0!important;padding:0!important;overflow:auto!important}`,
			);
			await loadDeck(page, SAMPLE_DECK, '/?fitPadding=0');
			const geometry = await settledGeometry(page);
			expectContained(geometry);
			expect(shortfall(geometry)).toBeLessThanOrEqual(1.5);
		});
	}

	test('a padding-only change of a content-box viewport refits the slide', async ({ page }) => {
		// The width is fixed, so changing the padding changes the client and border boxes but
		// not the content box a default ResizeObserver watches.
		await pin(
			page,
			{ width: 801.5, height: 700 },
			'[data-pptx-viewport]{box-sizing:content-box!important;border:0!important;padding:24px!important}',
		);
		await loadDeck(page, SAMPLE_DECK, '/?fitPadding=0');
		await settledGeometry(page);
		// In the injected style sheet, so it also wins over a binding's own inline padding.
		await page.evaluate(() => {
			const style = document.querySelector('#pinned-viewport')!;
			style.textContent = style.textContent!.replace('padding:24px', 'padding:0px');
		});
		await expect
			.poll(async () => (await settledGeometry(page)).slide.width, { timeout: 5000 })
			.toBeLessThanOrEqual(801.5 + LAYOUT_UNIT);
		const geometry = await settledGeometry(page);
		expect(geometry.content.width).toBeCloseTo(801.5, 1);
		expectContained(geometry);
		expect(shortfall(geometry)).toBeLessThan(0.1);
	});
});
