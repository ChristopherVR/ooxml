import { test, expect } from '@playwright/test';
import { openFind, taskPane } from './ribbon.js';
import { createVsdxFixture } from './fixture.mjs';

test('mobile Tools exposes Visio groups, opens Find and edits text with F2', async ({ page }) => {
	await page.setViewportSize({ width: 390, height: 844 });
	await page.goto('/demo/?sample=1');
	const viewer = page.locator('visio-viewer');
	await viewer.locator('[data-shape-id="s1"]').click();
	const tools = viewer.locator('.ribbon-tools>summary');
	await tools.click();
	for (const name of ['Pointer Tool', 'Find', 'Layers'])
		await expect(viewer.getByRole('button', { name, exact: true })).toBeVisible();
	await page.keyboard.press('Escape');
	await expect(tools).toBeFocused();
	const search = await openFind(viewer);
	await expect(viewer.locator('.ribbon-tools')).not.toHaveAttribute('open');
	await expect(search).toBeFocused();
	await search.fill('idea');
	await page.keyboard.press('Escape');
	await expect(search).toHaveValue('');
	await page.keyboard.press('Escape');
	await expect(viewer.locator('.find-bar')).toBeHidden();
	await page.keyboard.press('F2');
	await expect(viewer.locator('.inspector-pane')).toBeVisible();
	await expect(viewer.getByLabel('Selected shape text', { exact: true })).toBeVisible();
	await viewer.getByRole('button', { name: 'Close inspector' }).click();
});

test('mobile documentation menus keep the heading near the suite position and navigate sections', async ({
	page,
}) => {
	await page.setViewportSize({ width: 390, height: 844 });
	await page.goto('/docs/index.html');
	const title = await page.locator('.doc-content h1').boundingBox();
	expect(title!.y).toBeLessThan(170);
	await page.locator('.doc-mobile-menu>summary').click();
	await expect(page.getByRole('link', { name: 'Framework bindings', exact: true })).toBeVisible();
	await page.keyboard.press('Escape');
	await expect(page.locator('.doc-mobile-menu>summary')).toBeFocused();
	await page.locator('.doc-mobile-outline>summary').click();
	await page
		.locator('.doc-mobile-outline')
		.getByRole('link', { name: 'Try the local playground', exact: true })
		.click();
	await expect(page.locator('.doc-mobile-outline')).not.toHaveAttribute('open');
	await expect(page.locator('#playground')).toBeInViewport();
	await page.locator('.doc-mobile-menu>summary').click();
	await page
		.locator('.doc-mobile-menu')
		.getByRole('link', { name: 'Architecture', exact: true })
		.click();
	await expect(page.locator('.doc-content h1')).toContainText('Architecture');
});

test('opening screen supports browse cancellation, rejected input and repeated sample entry', async ({
	page,
}) => {
	await page.goto('/demo/');
	await expect(page.locator('#start-screen')).toBeVisible();
	const chooser = page.waitForEvent('filechooser');
	await page.getByRole('button', { name: 'Browse files', exact: true }).click();
	await chooser;
	await expect(page.locator('#start-screen')).toBeVisible();
	await page.locator('#file').setInputFiles({
		name: 'broken.vsdx',
		mimeType: 'application/octet-stream',
		buffer: Buffer.from('not a zip'),
	});
	await expect(page.locator('#error')).toBeVisible();
	await expect(page.locator('#start-screen')).toBeVisible();
	await page.getByRole('button', { name: 'Open the sample drawing', exact: true }).click();
	await expect(page.locator('#start-screen')).toBeHidden();
	await expect(page.locator('#error')).toBeHidden();
	await expect(page.locator('visio-viewer .viewport>svg')).toBeVisible();
	await page.locator('#file-menu').click();
	await expect(page.locator('#file-commands')).toHaveAttribute('data-open', '');
	await page.keyboard.press('Escape');
	await expect(page.locator('#file-menu')).toBeFocused();
	await expect(page.locator('#file-commands')).not.toHaveAttribute('data-open');
	await page.locator('#file-menu').click();
	await page.getByRole('button', { name: 'Load sample', exact: true }).click();
	await expect(page.locator('#file-commands')).not.toHaveAttribute('data-open');
});

test('mobile local file opens into bounded canvas and inspector can close and survive resize', async ({
	page,
}) => {
	await page.setViewportSize({ width: 390, height: 844 });
	await page.goto('/demo/');
	await page.locator('#file').setInputFiles({
		name: 'local.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture(),
	});
	await expect(page.locator('#start-screen')).toBeHidden();
	await expect(page.locator('#file-name')).toHaveText('local.vsdx');
	const viewer = page.locator('visio-viewer');
	await expect(viewer.locator('.inspector-pane')).toBeHidden();
	await taskPane(viewer, 'Inspector');
	await expect(viewer.locator('.inspector-pane')).toBeVisible();
	await viewer.getByRole('button', { name: 'Close inspector', exact: true }).click();
	await page.setViewportSize({ width: 1440, height: 900 });
	await expect(viewer.locator('.inspector-pane')).toBeHidden();
	await taskPane(viewer, 'Inspector');
	await page.setViewportSize({ width: 390, height: 844 });
	await expect(viewer.locator('.inspector-pane')).toBeVisible();
	await viewer.getByRole('button', { name: 'Close inspector', exact: true }).click();
	expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(390);
});
