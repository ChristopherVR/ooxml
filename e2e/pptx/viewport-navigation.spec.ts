/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Zoom at the pointer and pan a zoomed slide, the same in all five demos.
 *
 * Ctrl+wheel used to scale the slide around its top-left corner, so the point
 * under the pointer slid away, and a zoomed slide could only be moved with the
 * scrollbars or the wheel. The shared viewport controller now keeps the point
 * under the pointer and pans on a middle-button or Space drag, and on a plain
 * drag in a read-only view.
 *
 * The element that scrolls differs per binding (the `[data-pptx-viewport]`
 * element, or its `<main>` parent), so the spec finds it from the slide.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import type { PowerPointViewerAPI } from 'pptx-viewer-shared';

import { loadDeck, slideStage, zoomFitButton } from './support/deck';

declare global {
	interface Window {
		__pptxViewer: PowerPointViewerAPI;
	}
}

test.use({ viewport: { width: 1280, height: 800 } });

interface Scroll {
	left: number;
	top: number;
}

/** Scroll offsets of the element that scrolls the main slide. */
function scrollOf(page: Page): Promise<Scroll> {
	return page.evaluate(() => {
		const stage = document.querySelector('[data-pptx-viewport] [aria-roledescription="slide"]');
		for (let el = stage?.parentElement ?? null; el; el = el.parentElement) {
			if (/(auto|scroll)/.test(getComputedStyle(el).overflow)) {
				return { left: el.scrollLeft, top: el.scrollTop };
			}
		}
		return { left: Number.NaN, top: Number.NaN };
	});
}

/** Scroll the slide's scroller to `scroll`, so a pan in any direction has room. */
async function scrollTo(page: Page, scroll: Scroll): Promise<void> {
	await page.evaluate((target) => {
		const stage = document.querySelector('[data-pptx-viewport] [aria-roledescription="slide"]');
		for (let el = stage?.parentElement ?? null; el; el = el.parentElement) {
			if (/(auto|scroll)/.test(getComputedStyle(el).overflow)) {
				el.scrollLeft = target.left;
				el.scrollTop = target.top;
				return;
			}
		}
	}, scroll);
}

/** The main slide's laid-out width. */
async function slideWidth(page: Page): Promise<number> {
	return (await slideStage(page).boundingBox())!.width;
}

/** The client point at a fraction of the main slide. */
async function slidePoint(page: Page, fx: number, fy: number): Promise<{ x: number; y: number }> {
	const box = (await slideStage(page).boundingBox())!;
	return { x: box.x + fx * box.width, y: box.y + fy * box.height };
}

/** The fraction of the main slide under a client point. */
async function fractionAt(page: Page, x: number, y: number): Promise<{ fx: number; fy: number }> {
	const box = (await slideStage(page).boundingBox())!;
	return { fx: (x - box.x) / box.width, fy: (y - box.y) / box.height };
}

async function ctrlWheel(
	page: Page,
	x: number,
	y: number,
	deltaY: number,
	times: number,
): Promise<void> {
	await page.mouse.move(x, y);
	await page.keyboard.down('Control');
	for (let i = 0; i < times; i++) {
		await page.mouse.wheel(0, deltaY);
		await page.waitForTimeout(40);
	}
	await page.keyboard.up('Control');
}

/** Zoom in well past fit with the wheel at the slide's centre and wait for the layout. */
async function zoomIn(page: Page): Promise<void> {
	const before = (await slideStage(page).boundingBox())!.width;
	const centre = await slidePoint(page, 0.5, 0.5);
	await ctrlWheel(page, centre.x, centre.y, -100, 6);
	await expect
		.poll(async () => (await slideStage(page).boundingBox())!.width)
		.toBeGreaterThan(before * 1.5);
}

/** Positions of the active slide's elements, to show a drag moved none of them. */
function elementPositions(page: Page): Promise<string> {
	return page.evaluate(() =>
		JSON.stringify(
			window.__pptxViewer.getElements(0).map((element) => [element.id, element.x, element.y]),
		),
	);
}

test.describe('viewport navigation', () => {
	test.beforeEach(async ({ page }) => {
		await loadDeck(page);
		await expect(slideStage(page)).toBeVisible();
		await page.waitForFunction(() => typeof window.__pptxViewer?.getElements === 'function');
	});

	test('Ctrl+wheel keeps the slide point under the pointer', async ({ page }) => {
		const width = (await slideStage(page).boundingBox())!.width;
		const point = await slidePoint(page, 0.2, 0.25);
		await ctrlWheel(page, point.x, point.y, -100, 6);
		await expect
			.poll(async () => (await slideStage(page).boundingBox())!.width)
			.toBeGreaterThan(width * 1.5);
		await expect
			.poll(async () => {
				const { fx, fy } = await fractionAt(page, point.x, point.y);
				return Math.max(Math.abs(fx - 0.2), Math.abs(fy - 0.25));
			})
			.toBeLessThan(0.01);
	});

	test('a middle-button drag pans the zoomed slide', async ({ page }) => {
		await zoomIn(page);
		await scrollTo(page, { left: 300, top: 200 });
		const start = await scrollOf(page);
		const centre = await slidePoint(page, 0.5, 0.5);
		await page.mouse.move(centre.x, centre.y);
		await page.mouse.down({ button: 'middle' });
		await page.mouse.move(centre.x + 120, centre.y + 80, { steps: 6 });
		await page.mouse.up({ button: 'middle' });
		await expect
			.poll(() => scrollOf(page))
			.toEqual({ left: start.left - 120, top: start.top - 80 });
	});

	test('Space+drag pans without moving the shape under the pointer', async ({ page }) => {
		await zoomIn(page);
		await scrollTo(page, { left: 300, top: 200 });
		const start = await scrollOf(page);
		const positions = await elementPositions(page);
		const centre = await slidePoint(page, 0.5, 0.5);
		await page.mouse.move(centre.x, centre.y);
		await page.keyboard.down('Space');
		await page.mouse.down();
		await page.mouse.move(centre.x - 100, centre.y - 60, { steps: 6 });
		await page.mouse.up();
		await page.keyboard.up('Space');
		await expect
			.poll(() => scrollOf(page))
			.toEqual({ left: start.left + 100, top: start.top + 60 });
		expect(await elementPositions(page)).toBe(positions);
	});

	test('Space+drag pans with focus in the viewer, and stays on the slide', async ({ page }) => {
		await zoomIn(page);
		await scrollTo(page, { left: 300, top: 200 });
		const centre = await slidePoint(page, 0.5, 0.5);
		// A click puts focus in the viewer, where a bare Space also means "next slide".
		await page.mouse.click(centre.x, centre.y);
		const start = await scrollOf(page);
		await page.keyboard.down('Space');
		await page.mouse.down();
		await page.mouse.move(centre.x - 100, centre.y - 60, { steps: 6 });
		await page.mouse.up();
		await page.keyboard.up('Space');
		await expect
			.poll(() => scrollOf(page))
			.toEqual({ left: start.left + 100, top: start.top + 60 });
		expect(await page.evaluate(() => window.__pptxViewer.getActiveSlideIndex())).toBe(0);
	});

	test('Space presses a focused zoom button while the pointer is over the slide', async ({
		page,
	}) => {
		await zoomIn(page);
		const centre = await slidePoint(page, 0.5, 0.5);
		await page.mouse.move(centre.x, centre.y);
		const before = (await slideStage(page).boundingBox())!.width;
		await page
			.getByRole('button', { name: /^zoom in$/iu })
			.first()
			.focus();
		await page.keyboard.press('Space');
		await expect
			.poll(async () => (await slideStage(page).boundingBox())!.width)
			.toBeGreaterThan(before + 1);
	});

	test('a plain drag pans only in a read-only view', async ({ page }) => {
		await zoomIn(page);
		await scrollTo(page, { left: 300, top: 200 });
		const centre = await slidePoint(page, 0.5, 0.5);
		const drag = async () => {
			await page.mouse.move(centre.x, centre.y);
			await page.mouse.down();
			await page.mouse.move(centre.x - 90, centre.y - 50, { steps: 6 });
			await page.mouse.up();
		};

		await page.evaluate(() => window.__pptxViewer.setMode('preview'));
		await expect.poll(() => page.evaluate(() => window.__pptxViewer.getMode())).toBe('preview');
		await scrollTo(page, { left: 300, top: 200 });
		const start = await scrollOf(page);
		await drag();
		await expect.poll(() => scrollOf(page)).toEqual({ left: start.left + 90, top: start.top + 50 });

		// In the editor the same drag belongs to selection and moving shapes.
		await page.evaluate(() => window.__pptxViewer.setMode('edit'));
		await expect.poll(() => page.evaluate(() => window.__pptxViewer.getMode())).toBe('edit');
		await scrollTo(page, { left: 300, top: 200 });
		const editStart = await scrollOf(page);
		await drag();
		await page.waitForTimeout(200);
		expect(await scrollOf(page)).toEqual(editStart);
	});

	test('zooming out goes below fit, down to 10%', async ({ page }) => {
		const fitted = await slideWidth(page);
		const zoomOut = page.getByRole('button', { name: /^zoom out$/iu }).first();
		for (let i = 0; i < 5; i++) {
			await zoomOut.click();
		}
		await expect.poll(async () => (await slideWidth(page)) / fitted).toBeCloseTo(0.5, 2);

		const centre = await slidePoint(page, 0.5, 0.5);
		await ctrlWheel(page, centre.x, centre.y, 100, 10);
		await expect.poll(async () => (await slideWidth(page)) / fitted).toBeCloseTo(0.1, 2);

		// The host's setZoom goes below fit too. Vanilla's takes an absolute scale,
		// which for this deck and window is also below fit.
		await page.evaluate(() => window.__pptxViewer.setZoom(0.5));
		await expect.poll(() => page.evaluate(() => window.__pptxViewer.getZoom())).toBeCloseTo(0.5);
		expect(await slideWidth(page)).toBeLessThan(fitted - 50);
	});

	test('another slide keeps the zoom and is shown from its top-left corner', async ({ page }) => {
		const fitted = await slideWidth(page);
		await zoomIn(page);
		const zoomed = await slideWidth(page);
		await scrollTo(page, { left: 300, top: 200 });
		expect(await scrollOf(page)).toEqual({ left: 300, top: 200 });

		const second = page.locator('[aria-label="Go to slide 2"]').first();
		await second.click();
		await expect(second).toHaveAttribute('aria-current', /^(true|page)$/);
		await expect.poll(() => scrollOf(page)).toEqual({ left: 0, top: 0 });
		expect(Math.abs((await slideWidth(page)) - zoomed)).toBeLessThan(1);

		// Below fit, from the host, the zoom stays too.
		await page.evaluate(() => window.__pptxViewer.setZoom(0.5));
		await expect.poll(() => slideWidth(page)).toBeLessThan(fitted - 50);
		const small = await slideWidth(page);
		await page.evaluate(() => window.__pptxViewer.goTo(0));
		await expect.poll(() => page.evaluate(() => window.__pptxViewer.getActiveSlideIndex())).toBe(0);
		await page.waitForTimeout(300);
		expect(Math.abs((await slideWidth(page)) - small)).toBeLessThan(1);

		// Zoom to fit still lands on the same fit.
		await zoomFitButton(page).click();
		await expect.poll(async () => Math.abs((await slideWidth(page)) - fitted)).toBeLessThan(1);
	});
});
