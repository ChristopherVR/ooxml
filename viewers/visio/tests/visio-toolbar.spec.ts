import { test, expect } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';

test('Visio ribbon draws, deletes, undoes and redoes with tools and shortcuts', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1440, height: 1000 });
	await page.goto('/demo/?sample=1');
	const viewer = page.locator('visio-viewer');
	const ribbon = (name: string) =>
		viewer.locator('.toolbar').getByRole('button', { name, exact: true });
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
	await expect(viewer.locator('office-ui-menu-button[command="rectangle"]')).toHaveAttribute(
		'data-active',
		'',
	);
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

	// Visio's full Home layout is present; commands the core lacks are disabled with a reason.
	await expect(ribbon('Bold')).toBeDisabled();
	await expect(ribbon('Bold')).toHaveAttribute('title', /not available yet\. Needs core text/);
	// Dropdowns open as top-layer menus that stay inside the window, even at the right edge.
	await expect(viewer.getByRole('button', { name: 'Select', exact: true })).toHaveAttribute(
		'title',
		/not available yet\. Needs multi-shape selection/,
	);
	await viewer.getByRole('button', { name: 'Find', exact: true }).click();
	const menu = viewer.locator('office-ui-menu-button[data-menu="find"] >> [role="menu"]');
	await expect(menu).toBeVisible();
	const box = await menu.boundingBox();
	expect(box!.x + box!.width).toBeLessThanOrEqual(1440);
	await page.keyboard.press('Escape');
	await expect(menu).toBeHidden();

	// Visio's Shapes window: drag the Rectangle master onto the page to drop a shape there.
	const shapesBefore = await viewer.locator('svg.paper > g > [data-shape-id]').count();
	await viewer
		.locator('[data-master="rectangle"]')
		.first()
		.dragTo(viewer.locator('svg.paper'), { targetPosition: { x: 120, y: 120 } });
	await expect(viewer.locator('svg.paper > g > [data-shape-id]')).toHaveCount(shapesBefore + 1);
	await expect(viewer.locator('[data-status]')).toHaveText(/added from Basic Shapes/);
	await expect(viewer.locator('[data-master="circle"]').first()).toHaveAttribute(
		'aria-disabled',
		'true',
	);

	await viewer.getByRole('tab', { name: 'View', exact: true }).click();
	await viewer.locator('[data-check="ruler"]').click();
	await expect(viewer.locator('.canvas-area')).toHaveAttribute('data-ruler', 'true');
	await expect(viewer.locator('office-ui-ruler.ruler-h')).toBeVisible();
	await viewer.locator('[data-check="grid"]').click();
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
