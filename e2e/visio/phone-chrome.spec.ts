import { test, expect } from '@playwright/test';
import { openDemo } from './demo-page';

test.use({ hasTouch: true });

test('a host that stops rendering ends Presentation Mode, so nothing paints over the next tab', async ({
	page,
}) => {
	await page.setViewportSize({ width: 412, height: 915 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await expect(viewer.locator('svg.paper')).toHaveCount(1);
	await viewer.getByRole('button', { name: 'Presentation Mode' }).click();
	const overlay = viewer.locator('.presentation');
	await expect(overlay).toBeVisible();
	// The suite hides an inactive editor's host; the overlay (fixed, full screen) must go with it.
	await viewer.evaluate((node: HTMLElement) => {
		node.hidden = true;
		node.style.display = 'none';
	});
	await expect
		.poll(() =>
			viewer.evaluate((node) => ({
				hidden: node.shadowRoot!.querySelector<HTMLElement>('.presentation')!.hidden,
				full: node.shadowRoot!.fullscreenElement === null,
			})),
		)
		.toEqual({ hidden: true, full: true });
});

test('on a phone File is accent text and the groups chevron fades in instead of covering a group', async ({
	page,
}) => {
	await page.setViewportSize({ width: 412, height: 915 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	const file = viewer.locator('office-ui-ribbon').locator('.file');
	await expect(file).toBeVisible();
	// A pointer resting on File paints its hover tint, so park it away from the ribbon first.
	await page.mouse.move(200, 600);
	await expect
		.poll(() => file.evaluate((node) => getComputedStyle(node).backgroundColor))
		.toBe('rgba(0, 0, 0, 0)');
	const more = viewer.locator('#home-panel > .ribbon-scroll-more');
	await expect(more).toBeVisible();
	// The chevron sits at the panel's end, and a gradient fades the content into it.
	const [button, panel] = await Promise.all([
		more.boundingBox(),
		viewer.locator('#home-panel').boundingBox(),
	]);
	expect(button!.x + button!.width).toBeLessThanOrEqual(panel!.x + panel!.width + 1);
	expect(
		await more.evaluate((node) => getComputedStyle(node, '::before').backgroundImage),
	).toContain('linear-gradient');
});

test('the status bar keeps the zoom controls inside it on a tablet-width editor', async ({
	page,
}) => {
	await page.setViewportSize({ width: 768, height: 1024 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	const [bar, zoom] = await Promise.all([
		viewer.locator('office-ui-status-bar').boundingBox(),
		viewer.locator('office-ui-zoom-slider').boundingBox(),
	]);
	expect(zoom!.x + zoom!.width).toBeLessThanOrEqual(bar!.x + bar!.width + 1);
});
