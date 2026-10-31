import { expect, test } from '@playwright/test';
import { parseVsdx } from 'ooxml-core/visio';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { createVsdxFixture } from './fixture.mjs';
import { openDemo } from './demo-page';

test('Data Graphics are a group of callouts that read the shape data and follow its changes', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'Native.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Native target'),
	});
	await expect(viewer.locator('svg.paper')).toContainText('Native target');
	await viewer.getByRole('tab', { name: 'Data', exact: true }).click();
	const setData = (status: string, load: string) =>
		viewer.evaluate(
			async (node, values) => {
				const row = (name: string, type: 'string' | 'number', value: string) => ({
					type: 'set-shape-data' as const,
					pageId: '1',
					shapeId: '1',
					row: name,
					data: { label: name, prompt: '', type, format: '', value },
				});
				await (node as VisioViewerElement).applyEdits([
					row('Status', 'string', values.status),
					row('Load', 'number', values.load),
				]);
			},
			{ status, load },
		);
	await setData('Up', '60');
	const owner = async () => {
		const bytes = await viewer.evaluate((node) =>
			Array.from((node as VisioViewerElement).exportVsdx().bytes),
		);
		const shapes = (await parseVsdx(new Uint8Array(bytes))).pages[0]!.shapes;
		return shapes.find((item) => item.id === '1')!;
	};
	const percent = async () =>
		(await owner()).children[1]!.userCells?.find((cell) => cell.name === 'PercentFilled')?.value;
	await viewer.locator('svg.paper [data-shape-id="1"]').click();
	const graphics = viewer.locator('.data-graphics-dialog');
	await viewer.getByRole('button', { name: 'Data Graphics', exact: true }).click();
	await viewer.locator('office-ui-menu-item[command="graphic-text"]').click();
	await graphics.locator('[name="field"]').selectOption('Status');
	await graphics.locator('[label="Apply"] button').click();
	await expect(viewer.locator('svg.paper')).toContainText('Status: Up');
	await viewer.getByRole('button', { name: 'Data Graphics', exact: true }).click();
	await viewer.locator('office-ui-menu-item[command="graphic-bar"]').click();
	await graphics.locator('[name="field"]').selectOption('Load');
	await graphics.locator('[name="min"]').fill('0');
	await graphics.locator('[name="max"]').fill('100');
	await graphics.locator('[label="Apply"] button').click();
	await expect(viewer.locator('[data-status]')).toHaveText(/Applied data bar data graphics/i);
	// The shape is now a group of callout instances named by a data graphic master.
	const applied = await owner();
	expect(applied.kind).toBe('group');
	expect(applied.children.map((child) => child.name)).toEqual(['Text callout', 'Data bar 1']);
	expect(applied.userCells?.find((cell) => cell.name === 'visDGDisplayFormat')?.formula).toMatch(
		/^USE\(\{[0-9A-F-]{36}\}\)$/,
	);
	expect(await percent()).toBe('0.6');
	// A change to the shape's data reaches the callouts.
	await setData('Down', '25');
	await expect(viewer.locator('svg.paper')).toContainText('Status: Down');
	await expect(viewer.locator('svg.paper')).not.toContainText('Status: Up');
	expect(await percent()).toBe('0.25');
	// Position moves the callouts around the shape; they stay members of its group.
	await viewer.locator('svg.paper [data-shape-id="1"]').click();
	await viewer.getByRole('button', { name: 'Position', exact: true }).click();
	await viewer.locator('office-ui-menu-item[command="graphic-position-top"]').click();
	await expect(viewer.locator('[data-status]')).toHaveText(/Moved the data graphics/);
	await expect(async () => {
		const shape = await owner();
		expect(shape.children[0]!.rotation!.pinY).toBeGreaterThan(shape.height / 2);
	}).toPass();
	await page.screenshot({ path: test.info().outputPath('native-data-graphics.png') });
	// Remove Data Graphics takes the callouts off and leaves the shape.
	await viewer.getByRole('button', { name: 'Data Graphics', exact: true }).click();
	await viewer.locator('office-ui-menu-item[command="graphic-remove"]').click();
	await expect(viewer.locator('[data-status]')).toHaveText(/Removed data graphics/);
	const removed = await owner();
	expect(removed.kind).toBe('shape');
	expect(removed.children).toHaveLength(0);
	await expect(viewer.locator('svg.paper')).not.toContainText('Status: Down');
	expect(errors).toEqual([]);
});
