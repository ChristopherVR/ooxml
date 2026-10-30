import { test, expect, type Locator } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';
import { openDemo } from './demo-page';

async function menuItem(viewer: Locator, path: readonly string[]): Promise<void> {
	const effects = viewer.locator(
		'[data-ribbon-tab]:not([data-contextual]) office-ui-menu-button[data-menu="effects"]',
	);
	await effects.getByRole('button', { name: 'Effects', exact: true }).click();
	let scope: Locator = effects;
	for (const [index, name] of path.entries()) {
		const item = scope
			.getByRole('menuitem', { name, exact: true })
			.or(scope.getByRole('menuitemcheckbox', { name, exact: true }))
			.first();
		if (index === path.length - 1) {
			await item.click();
			return;
		}
		await item.hover();
		scope = scope.locator(`office-ui-menu-button[submenu][label="${name}"]`);
	}
}

test('perspective, extrusion and contour are drawn from the Effects menu and the Format Shape pane', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	await page.locator('#file').setInputFiles({
		name: 'effects-3d.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Solid me'),
	});
	await expect(page.locator('#file-name')).toHaveText('effects-3d.vsdx');
	const viewer = page.locator('visio-viewer');
	const shape = viewer.locator('svg.paper [data-shape-id="1"]');
	await shape.click();

	// An isometric preset has no perspective: still the cheap parallel matrix.
	await menuItem(viewer, ['3-D Rotation', 'Isometric Right Up']);
	await expect(shape.locator('[data-rotation3d]')).toHaveAttribute('transform', /^matrix\(/);

	// A perspective preset is projected point by point: no matrix, a polyline outline.
	await menuItem(viewer, ['3-D Rotation', 'Perspective Left']);
	await expect(shape.locator('[data-rotation3d]')).toHaveAttribute('data-solid3d', '');
	await expect(shape.locator('[data-rotation3d]')).not.toHaveAttribute('transform', /.+/);
	await expect(shape.locator('[data-rotation3d] [data-geometry]').first()).toHaveAttribute(
		'd',
		/^M [-\d. ]+ L /,
	);
	const flat = await shape.locator('[data-geometry]').first().getAttribute('d');

	// Depth needs a top bevel, as in Visio 16.
	await menuItem(viewer, ['Bevel', 'Circle']);
	await expect(shape.locator('[data-extrusion]')).toHaveCount(0);

	await menuItem(viewer, ['Glow', 'Glow Options...']);
	const pane = viewer.locator('.format-pane');
	await expect(pane).toBeVisible();
	const depth = pane.locator('[data-pane-field="extrusionDepth"]');
	await depth.fill('20');
	await depth.press('Tab');
	await expect(shape.locator('[data-solid] [data-extrusion]')).toHaveCount(1);
	await expect(shape.locator('[data-extrusion] path').first()).toBeAttached();
	const contour = pane.locator('[data-pane-field="contourSize"]');
	await contour.fill('4');
	await contour.press('Tab');
	await expect(shape.locator('[data-solid] [data-contour]')).toHaveCount(1);
	await expect(shape.locator('[data-geometry]').first()).toHaveAttribute('d', flat!);

	// 3-D Rotation > No Rotation removes the projection; depth and contour then show nothing.
	await menuItem(viewer, ['3-D Rotation', 'No Rotation']);
	await expect(shape.locator('[data-extrusion]')).toHaveCount(0);
	await expect(shape.locator('[data-rotation3d]')).toHaveCount(0);

	for (let step = 0; step < 6; step++) await viewer.locator('.qat [data-command="undo"]').click();
	await expect(shape.locator('[data-extrusion]')).toHaveCount(0);
	expect(errors).toEqual([]);
});
