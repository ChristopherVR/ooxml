import { test, expect } from '@playwright/test';

test('shared editor matches compact chrome geometry and keeps all navigation functional', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1440, height: 1000 });
	await page.goto('/demo/');
	const viewer = page.locator('visio-viewer');
	await expect(viewer.locator('svg')).toHaveCount(1);
	await expect(viewer.getByRole('tab', { name: 'Home', exact: true })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	const rail = await viewer.locator('.page-rail').boundingBox();
	const inspector = await viewer.locator('.inspector-pane').boundingBox();
	expect(rail?.width).toBe(180);
	expect(inspector?.width).toBe(288);
	expect((await viewer.locator('.ribbon-primary').boundingBox())?.height).toBe(32);
	expect((await viewer.locator('.ribbon-tabs').boundingBox())?.height).toBe(35);
	expect((await viewer.locator('.status').boundingBox())?.height).toBe(29);
	await viewer.getByRole('button', { name: 'Go to page 2: Architecture' }).click();
	await expect(viewer.locator('svg')).toHaveAttribute('aria-label', 'Architecture');
	await expect(viewer.getByRole('combobox', { name: 'Page', exact: true })).toHaveValue('1');
	await expect(viewer.locator('[data-page-index="1"]')).toHaveAttribute('aria-current', 'page');
	await viewer.getByRole('tab', { name: 'View', exact: true }).click();
	await viewer.getByRole('button', { name: 'Pages pane' }).click();
	await expect(viewer.locator('.page-rail')).toBeHidden();
	await viewer.getByRole('button', { name: 'Pages pane' }).click();
	await expect(viewer.locator('.page-rail')).toBeVisible();
	await viewer.getByRole('button', { name: 'Zoom in', exact: true }).click();
	await viewer.getByRole('button', { name: '100%', exact: true }).click();
	await expect(viewer.locator('output')).toHaveText('100%');
	await viewer.getByRole('button', { name: 'Fit page', exact: true }).click();
	const canvas = await viewer.locator('.viewport').boundingBox();
	const drawing = await viewer.locator('.viewport > svg').boundingBox();
	expect(
		Math.abs(drawing!.y + drawing!.height / 2 - canvas!.y - canvas!.height / 2),
	).toBeLessThanOrEqual(1);
	await viewer.getByRole('tab', { name: 'View', exact: true }).focus();
	await page.keyboard.press('ArrowLeft');
	await expect(viewer.getByRole('tab', { name: 'Home', exact: true })).toBeFocused();
	await viewer.getByRole('searchbox', { name: 'Search diagram text' }).fill('framework');
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
	await page.goto('/demo/');
	const viewer = page.locator('visio-viewer');
	await expect(viewer.locator('.page-rail')).toBeHidden();
	await expect(viewer.locator('svg')).toBeVisible();
	await viewer.getByRole('combobox', { name: 'Page', exact: true }).selectOption('1');
	await expect(viewer.locator('svg')).toHaveAttribute('aria-label', 'Architecture');
	for (const control of [
		viewer.getByRole('combobox', { name: 'Page', exact: true }),
		viewer.getByRole('button', { name: 'Fit page', exact: true }),
		viewer.getByRole('button', { name: 'Zoom in', exact: true }),
		viewer.locator('.edit-controls summary'),
	]) {
		const box = await control.boundingBox();
		expect(box?.width).toBeGreaterThanOrEqual(44);
		expect(box?.height).toBeGreaterThanOrEqual(44);
	}
	await viewer.getByRole('tab', { name: 'View', exact: true }).click();
	await viewer.getByRole('button', { name: 'Pages pane' }).click();
	await expect(viewer.locator('.page-rail')).toBeVisible();
	await viewer.locator('[data-page-index="0"]').click();
	await expect(viewer.locator('svg')).toHaveAttribute('aria-label', 'Release workflow');
	await viewer.getByRole('button', { name: 'Pages pane' }).click();
	await viewer.locator('.edit-controls summary').click();
	await expect(viewer.getByLabel('Selected shape text', { exact: true })).toBeVisible();
	expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
});

test.describe('touch-enabled tablet chrome', () => {
	test.use({ hasTouch: true, viewport: { width: 1024, height: 900 } });
	test('enlarges compact commands to touch targets for coarse pointers', async ({ page }) => {
		await page.goto('/demo/');
		const viewer = page.locator('visio-viewer');
		await expect(viewer.locator('svg')).toBeVisible();
		for (const selector of [
			'[data-tab="home"]',
			'[data-chrome="next-page"]',
			'[data-action="fit"]',
			'[data-action="in"]',
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
