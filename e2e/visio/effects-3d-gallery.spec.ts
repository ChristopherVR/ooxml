import { test, expect, type Locator } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';
import { openDemo } from './demo-page';

async function openEffects(viewer: Locator): Promise<Locator> {
	const effects = viewer.locator(
		'[data-ribbon-tab]:not([data-contextual]) office-ui-menu-button[data-menu="effects"]',
	);
	await effects.getByRole('button', { name: 'Effects', exact: true }).click();
	return effects;
}
async function menuItem(viewer: Locator, path: readonly string[]): Promise<void> {
	const effects = await openEffects(viewer);
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

test('the 3-D Rotation gallery follows the Office presets; bevels are lit from a profile; text follows perspective', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	await page.locator('#file').setInputFiles({
		name: 'effects-3d-gallery.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Gallery me'),
	});
	await expect(page.locator('#file-name')).toHaveText('effects-3d-gallery.vsdx');
	const viewer = page.locator('visio-viewer');
	const shape = viewer.locator('svg.paper [data-shape-id="1"]');
	await shape.click();

	// The gallery lists the Office presets, in order.
	const effects = await openEffects(viewer);
	await effects.getByRole('menuitem', { name: '3-D Rotation', exact: true }).hover();
	for (const name of [
		'Isometric Bottom Down',
		'Off Axis 1 Top',
		'Off Axis 2 Left',
		'Perspective Relaxed Moderately',
		'Perspective Contrasting Left',
		'Perspective Heroic Extreme Right',
	])
		await expect(
			effects.getByRole('menuitemcheckbox', { name, exact: true }).first(),
		).toBeAttached();
	await page.keyboard.press('Escape');

	// A parallel preset is one matrix; a perspective preset fits the text with its own matrix.
	await menuItem(viewer, ['3-D Rotation', 'Isometric Top Up']);
	await expect(shape.locator('[data-rotation3d]')).toHaveAttribute('transform', /^matrix\(/);
	await menuItem(viewer, ['Bevel', 'Angle']);
	await menuItem(viewer, ['3-D Rotation', 'Perspective Relaxed']);
	await expect(shape.locator('[data-rotation3d]')).toHaveAttribute('data-solid3d', '');
	await expect(shape.locator('[data-solid-flat]')).toHaveAttribute('transform', /^matrix\(/);
	await expect(shape.locator('[data-solid-flat] text').first()).toBeAttached();

	// The bevel is a profile through a height-field filter: a table of heights and one light.
	await menuItem(viewer, ['3-D Rotation', 'No Rotation']);
	await expect(shape.locator('[data-bevel]')).toHaveAttribute('data-bevel-type', '5');
	const angle = await viewer
		.locator('svg.paper filter feComponentTransfer feFuncA')
		.first()
		.getAttribute('tableValues');
	expect(angle!.split(' ')).toHaveLength(128);
	await menuItem(viewer, ['Bevel', 'Circle']);
	await expect(shape.locator('[data-bevel]')).toHaveAttribute('data-bevel-type', '1');
	const circle = await viewer
		.locator('svg.paper filter feComponentTransfer feFuncA')
		.last()
		.getAttribute('tableValues');
	expect(circle).not.toBe(angle);
	await expect(viewer.locator('svg.paper filter feDistantLight').last()).toHaveAttribute(
		'azimuth',
		/^-?[\d.]+$/,
	);
	expect(errors).toEqual([]);
});
