import { test, expect } from '@playwright/test';
import { openDemo } from './demo-page';
import { createVsdxFixture } from './fixture.mjs';
import { history, ribbonGroup } from './ribbon';

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

test('on a phone the Illustrations group inserts an icon', async ({ page }) => {
	await page.setViewportSize({ width: 375, height: 812 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	const paper = viewer.locator('svg.paper');
	await expect(paper).toContainText('Release workflow');
	const images = paper.locator('image');
	const before = await images.count();
	await viewer.getByRole('tab', { name: 'Insert', exact: true }).tap();
	await ribbonGroup(viewer, 'Illustrations');
	await viewer.getByRole('button', { name: 'Icons', exact: true }).tap();
	await expect(page.getByRole('dialog', { name: 'Icons' })).toBeVisible();
	await page.getByRole('searchbox', { name: 'Search icons' }).fill('Light bulb');
	await page.getByRole('button', { name: 'Light bulb', exact: true }).tap();
	await viewer
		.locator('office-ui-dialog[open]')
		.getByRole('button', { name: 'Insert', exact: true })
		.tap();
	await expect(images).toHaveCount(before + 1);
});

test.describe('at narrow phone width', () => {
	test.use({ viewport: { width: 320, height: 720 }, isMobile: true });

	test('finds diagram text and uses nested View commands after paging the tabs', async ({
		page,
	}) => {
		await openDemo(page);
		const viewer = page.locator('visio-viewer');
		const home = viewer.locator('#home-panel');
		await home.getByRole('button', { name: 'More', exact: true }).tap();
		await expect.poll(() => home.evaluate((panel) => panel.scrollLeft)).toBeGreaterThan(0);
		await home.getByRole('button', { name: 'Editing', exact: true }).tap();
		const find = viewer.getByRole('button', { name: 'Find', exact: true });
		await find.tap();
		const menu = viewer.locator('office-ui-menu-button[data-menu="find"]');
		const findText = menu.getByText('Find...', { exact: true });
		await expect(findText).toBeVisible();
		await page.keyboard.press('Escape');
		await expect(findText).toBeHidden();
		await expect(find).toBeVisible();
		await find.tap();
		await findText.tap();
		const query = viewer.getByRole('searchbox', { name: 'Search diagram text' });
		await query.fill('Start with an idea');
		await expect(viewer.locator('office-ui-find-bar [role="status"]')).toHaveText(
			'1 matching shapes',
		);

		const ribbon = viewer.locator('office-ui-ribbon');
		await ribbon.getByRole('button', { name: 'More tabs', exact: true }).tap();
		await expect(ribbon).toHaveAttribute('data-tabs-start', '');
		await ribbon.getByRole('tab', { name: 'View', exact: true }).tap();
		await ribbonGroup(viewer, 'Zoom');
		const zoom = viewer.locator('office-ui-menu-button[data-menu="zoom"]');
		await zoom.getByRole('button', { name: 'Zoom', exact: true }).tap();
		await zoom.locator('office-ui-menu-item[label="200%"]').tap();
		await expect(viewer.locator('office-ui-zoom-slider output')).toHaveText('200%');
	});

	test('formats a shape and uses the phone Undo, Redo and Save controls', async ({ page }) => {
		await openDemo(page);
		await page.locator('#file').setInputFiles({
			name: 'phone-toolbar.vsdx',
			mimeType: 'application/vnd.ms-visio.drawing',
			buffer: await createVsdxFixture('Phone diagram'),
		});
		const viewer = page.locator('visio-viewer');
		const shape = viewer.locator('svg.paper [data-shape-id="1"]');
		await expect(shape).toHaveAttribute('aria-label', 'Phone diagram');
		await shape.tap();
		await ribbonGroup(viewer, 'Font');
		await viewer.getByRole('button', { name: 'Bold', exact: true }).tap();
		const textRun = shape.locator('tspan[font-family]').first();
		await expect(textRun).toHaveAttribute('font-weight', 'bold');
		await history(viewer, 'Undo').tap();
		await expect(textRun).not.toHaveAttribute('font-weight', 'bold');
		await history(viewer, 'Redo').tap();
		await expect(textRun).toHaveAttribute('font-weight', 'bold');
		const download = page.waitForEvent('download');
		await viewer.locator('.qat').getByRole('button', { name: 'Save', exact: true }).tap();
		expect((await download).suggestedFilename()).toMatch(/\.vsdx$/);
	});
});
