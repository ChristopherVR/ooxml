import { expect, test, type Page } from '@playwright/test';
import { cell, fixture, shape, rectangle } from '../../src/core/visio/test-fixtures';
import { openDemo } from './demo-page';

/** The `data-shape-id` of the focused shape, through the viewer's shadow root. */
const focusedShape = (page: Page) =>
	page.evaluate(() => {
		let active = document.activeElement;
		while (active?.shadowRoot?.activeElement) active = active.shadowRoot.activeElement;
		return active?.closest?.('[data-shape-id]')?.getAttribute('data-shape-id') ?? null;
	});

/** A shape with a size and position, so an arrow key has something to move. */
const PLACED =
	cell('PinX', 3) +
	cell('PinY', 3) +
	cell('Width', 2) +
	cell('Height', 1) +
	cell('LocPinX', 1) +
	cell('LocPinY', 0.5);

async function openDrawing(page: Page) {
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const bytes = await fixture({
		pages: [{ id: '0', contents: `<Shapes>${shape('1', PLACED + rectangle)}</Shapes>` }],
	});
	await page.locator('#file').setInputFiles({
		name: 'keyboard.vsdx',
		mimeType: 'application/octet-stream',
		buffer: Buffer.from(bytes),
	});
	await expect(page.locator('#file-name')).toHaveText('keyboard.vsdx');
	return page.locator('visio-viewer');
}

test('a shape keeps keyboard focus through nudges, and Delete still works', async ({ page }) => {
	const viewer = await openDrawing(page);
	const target = viewer.locator('.paper [data-shape-id="1"]');
	await target.focus();
	await page.keyboard.press('Enter');
	await expect(target).toHaveAttribute('aria-pressed', 'true');
	const left = async () =>
		(await viewer.locator('.paper [data-shape-id="1"]').boundingBox())?.x ?? Number.NaN;

	for (let nudge = 1; nudge <= 3; nudge += 1) {
		const before = await left();
		await page.keyboard.press('ArrowRight');
		await expect.poll(left).toBeGreaterThan(before);
		expect(await focusedShape(page), `focus after nudge ${nudge}`).toBe('1');
	}

	// A key pressed while the last nudge is still saving is dropped, so Delete is repeated.
	await expect(async () => {
		await page.keyboard.press('Delete');
		await expect(viewer.locator('.paper [data-shape-id="1"]')).toHaveCount(0, { timeout: 500 });
	}).toPass();
	// With the shape gone the canvas holds focus, so the keyboard is not lost.
	await expect(viewer.locator('.viewport')).toBeFocused();
});

test('keyboard focus on a shape is visible', async ({ page }) => {
	const viewer = await openDrawing(page);
	const target = viewer.locator('.paper [data-shape-id="1"]');
	await target.focus();
	await page.keyboard.press('Shift+Tab');
	await page.keyboard.press('Tab');
	await expect(target).toBeFocused();
	const geometry = target.locator('[data-geometry]').first();
	await expect(geometry).toHaveCSS('stroke-dasharray', /\d/);
});
