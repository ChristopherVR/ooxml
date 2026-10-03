import { test, expect } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';

test('Visio ribbon draws, deletes, undoes and redoes with tools and shortcuts', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1440, height: 1000 });
	await page.goto('/demo/?sample=1');
	const viewer = page.locator('visio-viewer');
	const ribbon = (name: string) => viewer.getByRole('button', { name, exact: true });
	await expect(ribbon('Rectangle')).toBeDisabled();
	await page.locator('#file').setInputFiles({
		name: 'toolbar.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Existing shape'),
	});
	await expect(page.locator('#file-name')).toHaveText('toolbar.vsdx');
	await expect(ribbon('Undo')).toBeDisabled();
	await expect(ribbon('Undo')).toHaveAttribute('aria-keyshortcuts', 'Control+Z');

	await ribbon('Rectangle').click();
	await expect(ribbon('Rectangle')).toHaveAttribute('aria-pressed', 'true');
	const paper = await viewer.locator('svg.paper').boundingBox();
	const start = { x: paper!.x + paper!.width * 0.2, y: paper!.y + paper!.height * 0.6 };
	await page.mouse.move(start.x, start.y);
	await page.mouse.down();
	await page.mouse.move(start.x + 120, start.y + 60, { steps: 4 });
	await expect(viewer.locator('.draw-preview')).toHaveCount(1);
	await page.mouse.up();
	const created = viewer.locator('svg.paper [data-shape-id="2"]');
	await expect(created).toHaveCount(1);
	await expect(viewer.locator('.draw-preview')).toHaveCount(0);
	await expect(viewer.locator('[data-shape-status]')).toHaveAttribute(
		'value',
		/^Width: .* in {2}Height: .* in$/,
	);
	await expect(page.locator('#file-state')).toContainText('Edited copy');

	await page.keyboard.press('Escape');
	await expect(ribbon('Pointer Tool')).toHaveAttribute('aria-pressed', 'true');
	await created.click();
	// SVG outlines scale with drawing units; a focus ring would cover the page.
	expect(await created.evaluate((shape) => getComputedStyle(shape).outlineStyle)).toBe('none');
	await page.keyboard.press('Delete');
	await expect(created).toHaveCount(0);
	await page.keyboard.press('Control+z');
	await expect(created).toHaveCount(1);
	await page.keyboard.press('Control+y');
	await expect(created).toHaveCount(0);
	await ribbon('Undo').click();
	await expect(created).toHaveCount(1);

	await viewer.getByRole('tab', { name: 'View', exact: true }).click();
	await ribbon('Grid').click();
	await expect(viewer.locator('.viewport')).toHaveAttribute('data-grid', 'true');
	const before = await viewer.evaluate(
		(element) => (element as HTMLElement & { zoom: number }).zoom,
	);
	await ribbon('Page Width').click();
	const after = await viewer.evaluate(
		(element) => (element as HTMLElement & { zoom: number }).zoom,
	);
	expect(after).toBeGreaterThan(before);
	const canvas = viewer.locator('.viewport');
	expect(await canvas.evaluate((element) => element.scrollWidth <= element.clientWidth)).toBe(true);
});
