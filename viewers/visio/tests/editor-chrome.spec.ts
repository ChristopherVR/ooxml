import { test, expect } from '@playwright/test';
import { openFind, taskPane, zoomPreset } from './ribbon.js';

test('shared editor matches compact chrome geometry and keeps all navigation functional', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1440, height: 1000 });
	await page.goto('/demo/?sample=1');
	const viewer = page.locator('visio-viewer');
	await expect(viewer.locator('svg.paper')).toHaveCount(1);
	await expect(viewer.getByRole('tab', { name: 'Home', exact: true })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	// Visio's layout: Shapes window on the left, pages as bottom tabs, no page rail by default.
	const shapes = await viewer.locator('.shapes-pane').boundingBox();
	const inspector = await viewer.locator('.inspector-pane').boundingBox();
	expect(shapes?.width).toBe(232);
	await expect(viewer.locator('.page-rail')).toBeHidden();
	expect(inspector?.width).toBe(288);
	const tabs = await viewer.locator('office-ui-tab-strip').boundingBox();
	expect(tabs?.height).toBeGreaterThanOrEqual(26);
	expect(tabs?.height).toBeLessThanOrEqual(36);
	expect((await viewer.locator('.ribbon-tabs').boundingBox())?.height).toBe(36);
	expect((await viewer.locator('.status').boundingBox())?.height).toBeGreaterThanOrEqual(28);
	expect((await viewer.locator('.status').boundingBox())?.height).toBeLessThanOrEqual(36);
	await taskPane(viewer, 'Pages');
	await expect(viewer.locator('.shapes-pane')).toBeHidden();
	expect((await viewer.locator('.page-rail').boundingBox())?.width).toBe(180);
	await viewer.getByRole('button', { name: 'Go to page 2: Architecture' }).click();
	await expect(viewer.locator('svg.paper')).toHaveAttribute('aria-label', 'Architecture');
	await expect(viewer.getByRole('tab', { name: 'Architecture', exact: true })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	await expect(viewer.locator('[data-page-status]')).toHaveAttribute('value', 'Page 2 of 2');
	await expect(viewer.locator('[data-page-index="1"]')).toHaveAttribute('aria-current', 'page');
	await taskPane(viewer, 'Pages');
	await expect(viewer.locator('.page-rail')).toBeHidden();
	await taskPane(viewer, 'Shapes');
	await expect(viewer.locator('.shapes-pane')).toBeVisible();
	await viewer.getByRole('button', { name: 'Zoom in', exact: true }).click();
	await zoomPreset(viewer, 100);
	await expect(viewer.locator('output')).toHaveText('100%');
	await viewer.getByRole('button', { name: 'Fit to Window', exact: true }).click();
	const canvas = await viewer.locator('.viewport').boundingBox();
	const drawing = await viewer.locator('.viewport > svg').boundingBox();
	expect(
		Math.abs(drawing!.y + drawing!.height / 2 - canvas!.y - canvas!.height / 2),
	).toBeLessThanOrEqual(1);
	await viewer.getByRole('tab', { name: 'View', exact: true }).focus();
	await page.keyboard.press('ArrowLeft');
	await expect(viewer.getByRole('tab', { name: 'Review', exact: true })).toBeFocused();
	await page.keyboard.press('Home');
	await expect(viewer.getByRole('tab', { name: 'Home', exact: true })).toBeFocused();
	await (await openFind(viewer)).fill('framework');
	await viewer.getByRole('button', { name: 'Next matching shape' }).click();
	await expect(viewer.locator('.shape-inspector')).toBeVisible();
	await expect(viewer.locator('.shape-inspector')).toContainText('Your framework');
	await viewer.locator('.notes-strip button').click();
	await expect(viewer.locator('.notes')).toHaveAttribute('open', '');
	await expect(viewer.locator('.notes summary')).toBeFocused();
	expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(1440);
});

test('mobile keeps page navigation, editing disclosures and zoom controls reachable', async ({
	page,
}) => {
	await page.setViewportSize({ width: 390, height: 844 });
	await page.goto('/demo/?sample=1');
	const viewer = page.locator('visio-viewer');
	await expect(viewer.locator('.page-rail')).toBeHidden();
	await expect(viewer.locator('svg.paper')).toBeVisible();
	await viewer.getByRole('tab', { name: 'Home', exact: true }).click();
	await viewer.getByRole('tab', { name: 'Architecture', exact: true }).click();
	await expect(viewer.locator('svg.paper')).toHaveAttribute('aria-label', 'Architecture');
	const pagePicker = await viewer
		.getByRole('tab', { name: 'Architecture', exact: true })
		.boundingBox();
	expect(pagePicker?.width).toBeGreaterThanOrEqual(44);
	expect(pagePicker?.height).toBeGreaterThanOrEqual(44);
	await taskPane(viewer, 'Inspector');
	for (const control of [
		viewer.getByRole('button', { name: 'Fit page to current window', exact: true }),
		viewer.getByRole('button', { name: 'Zoom in', exact: true }),
		viewer.getByRole('button', { name: 'Next page', exact: true }),
		viewer.locator('.edit-controls summary'),
	]) {
		const box = await control.boundingBox();
		expect(box?.width).toBeGreaterThanOrEqual(44);
		expect(box?.height).toBeGreaterThanOrEqual(44);
	}
	await taskPane(viewer, 'Pages');
	await expect(viewer.locator('.page-rail')).toBeVisible();
	await viewer.locator('[data-page-index="0"]').click();
	await expect(viewer.locator('svg.paper')).toHaveAttribute('aria-label', 'Release workflow');
	await taskPane(viewer, 'Pages');
	await taskPane(viewer, 'Inspector');
	await viewer.locator('.edit-controls summary').click();
	await expect(viewer.getByLabel('Selected shape text', { exact: true })).toBeVisible();
	expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
});

test.describe('touch-enabled tablet chrome', () => {
	test.use({ hasTouch: true, viewport: { width: 1024, height: 900 } });
	test('enlarges compact commands to touch targets for coarse pointers', async ({ page }) => {
		await page.goto('/demo/?sample=1');
		const viewer = page.locator('visio-viewer');
		await expect(viewer.locator('svg.paper')).toBeVisible();
		for (const selector of [
			'[data-tab="home"]',
			'office-ui-tab-strip [aria-label="Next page"]',
			'office-ui-zoom-slider .fit',
			'office-ui-zoom-slider [aria-label="Zoom in"]',
			'.notes-strip button',
		]) {
			const box = await viewer.locator(selector).boundingBox();
			expect(box?.width).toBeGreaterThanOrEqual(44);
			expect(box?.height).toBeGreaterThanOrEqual(44);
		}
		expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(
			1024,
		);
	});
});
