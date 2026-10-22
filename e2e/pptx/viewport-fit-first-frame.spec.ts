/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The first animation frame on which the main slide exists must already show it at the fitted
 * scale. React used to measure its viewport in a passive effect, so it painted the slide at its
 * authored size for a frame and then jumped to the fitted size.
 *
 * The viewport is pinned to each size with injected CSS, so the same steps run unchanged
 * against all five demos.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { loadDeck, SAMPLE_DECK } from './support/deck';

test.use({ viewport: { width: 1600, height: 1000 } });

/** The sample deck's authored slide width (12192000 EMU). */
const AUTHORED_WIDTH = 1280;

interface Size {
	width: number;
	height: number;
}

declare global {
	interface Window {
		__firstPaintedSlide?: Size;
	}
}

/**
 * Pin the viewport to `size` (when given) and record the main slide's size as the browser is
 * about to paint the first frame that contains it.
 *
 * An animation frame callback runs before that frame's layout and resize observers, and a
 * binding may still correct the size in a resize observer before the paint. So the recorder
 * starts observing the slide from the first animation frame that finds it, and reads the size
 * in the observer's first callback: the last point before that frame is painted.
 */
async function instrument(page: Page, size?: Size): Promise<void> {
	await page.addInitScript((pinned) => {
		if (pinned) {
			const style = document.createElement('style');
			style.textContent = `[data-pptx-viewport]{flex:none!important;width:${pinned.width}px!important;height:${pinned.height}px!important;min-width:0!important;min-height:0!important;max-width:none!important;max-height:none!important}`;
			const add = () => document.head.append(style);
			if (document.head) {
				add();
			} else {
				document.addEventListener('DOMContentLoaded', add);
			}
		}
		const selector = '[data-pptx-viewport] [aria-roledescription="slide"]';
		const tick = () => {
			const slide = document.querySelector(selector);
			if (!slide || slide.getBoundingClientRect().width === 0) {
				requestAnimationFrame(tick);
				return;
			}
			const observer = new ResizeObserver(() => {
				const rect = document.querySelector(selector)?.getBoundingClientRect();
				if (rect) {
					window.__firstPaintedSlide = { width: rect.width, height: rect.height };
				}
				observer.disconnect();
			});
			observer.observe(slide);
		};
		requestAnimationFrame(tick);
	}, size);
}

async function settledSlide(page: Page) {
	await expect(
		page.locator('[data-pptx-viewport] [data-pptx-element="true"]').first(),
	).toBeVisible();
	// Let any late measurement land before reading the settled size.
	await page.evaluate(
		() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
	);
	return page.evaluate(() => {
		const slide = document.querySelector('[data-pptx-viewport] [aria-roledescription="slide"]')!;
		const rect = slide.getBoundingClientRect();
		return {
			slide: { width: rect.width, height: rect.height },
			firstPainted: window.__firstPaintedSlide ?? null,
		};
	});
}

test.describe('viewport fit', () => {
	test('the first painted frame already shows the fitted slide', async ({ page }) => {
		await instrument(page);
		await loadDeck(page, SAMPLE_DECK);
		const { slide, firstPainted } = await settledSlide(page);
		// The demo's viewport is narrower than the authored slide, so a fit is visible.
		expect(slide.width).toBeLessThan(AUTHORED_WIDTH);
		expect(firstPainted?.width).toBeCloseTo(slide.width, 1);
	});

	for (const [width, height] of [
		[1000, 700],
		[640, 480],
	] as const) {
		test(`the first painted frame is fitted in a ${width} x ${height} viewport`, async ({
			page,
		}) => {
			await instrument(page, { width, height });
			await loadDeck(page, SAMPLE_DECK);
			const { slide, firstPainted } = await settledSlide(page);
			expect(slide.width).toBeLessThan(width);
			expect(firstPainted?.width).toBeCloseTo(slide.width, 1);
			expect(firstPainted?.height).toBeCloseTo(slide.height, 1);
		});
	}
});
