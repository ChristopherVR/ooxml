import { expect, test } from '@playwright/test';
import { fixture, shape, cell } from '../../src/core/visio/test-fixtures';
import { openDemo } from './demo-page';
import { readFileSync } from 'node:fs';

test('native cached jumps update on the canvas and in saved geometry', async ({ page }) => {
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	await page.locator('#file').setInputFiles({
		name: 'native-jumps.vsdx',
		mimeType: 'application/octet-stream',
		buffer: readFileSync(
			new URL('../../src/core/visio/__fixtures__/native-cached-line-jumps.vsdx', import.meta.url),
		),
	});
	await expect(page.locator('#file-name')).toHaveText('native-jumps.vsdx');
	const viewer = page.locator('visio-viewer');
	const outline = viewer.locator('svg.paper [data-shape-id="1"] path').first();
	await expect(outline).toHaveAttribute('d', / A /);
	await viewer.getByRole('tab', { name: 'Design', exact: true }).click();
	await viewer.locator('office-ui-ribbon-group[launcher="page-setup-dialog"] .launcher').click();
	const dialog = viewer.locator('.page-setup-dialog');
	await dialog.getByRole('tab', { name: 'Layout and Routing', exact: true }).click();
	await dialog.getByRole('combobox', { name: 'Add line jumps to', exact: true }).click();
	await dialog.getByRole('option', { name: 'None', exact: true }).click();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(outline).not.toHaveAttribute('d', / A /);
	await viewer.locator('.qat [data-command="undo"]').click();
	await expect(outline).toHaveAttribute('d', / A /);
	await viewer.locator('.qat [data-command="redo"]').click();
	await expect(outline).not.toHaveAttribute('d', / A /);
	await viewer
		.locator('svg.paper')
		.screenshot({ path: test.info().outputPath('native-jumps-disabled.png') });
});

test('Layout and Routing edits native jump settings, undo and printed gaps', async ({ page }) => {
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const line = (id: string, x1: number, y1: number, x2: number, y2: number) =>
		shape(
			id,
			cell('BeginX', x1) +
				cell('BeginY', y1) +
				cell('EndX', x2) +
				cell('EndY', y2) +
				cell('ShapeRouteStyle', 16),
		);
	const bytes = await fixture({
		pages: [
			{
				id: '0',
				pageCells: cell('LineJumpCode', 0),
				contents: `<Shapes>${line('1', 1, 2, 5, 2)}${line('2', 3, 0, 3, 4)}</Shapes>`,
			},
		],
	});
	await page.locator('#file').setInputFiles({
		name: 'crossing-lines.vsdx',
		mimeType: 'application/octet-stream',
		buffer: Buffer.from(bytes),
	});
	const viewer = page.locator('visio-viewer');
	await expect(viewer.locator('svg.paper')).toBeVisible();
	await expect(page.locator('#file-name')).toHaveText('crossing-lines.vsdx');
	await viewer.getByRole('tab', { name: 'Design', exact: true }).click();
	await viewer.locator('office-ui-ribbon-group[launcher="page-setup-dialog"] .launcher').click();
	const dialog = viewer.locator('.page-setup-dialog');
	await dialog.getByRole('tab', { name: 'Layout and Routing', exact: true }).click();
	await expect(dialog.getByRole('tab', { name: 'Layout and Routing' })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	for (const [name, option] of [
		['Add line jumps to', 'Horizontal lines'],
		['Line jump style', 'Gap'],
		['Horizontal jump direction', 'Down'],
		['Vertical jump direction', 'Left'],
	]) {
		await dialog.getByRole('combobox', { name: name!, exact: true }).click();
		await dialog.getByRole('option', { name: option!, exact: true }).click();
	}
	await dialog
		.getByRole('spinbutton', { name: 'Horizontal jump size (%)', exact: true })
		.fill('50');
	await dialog.getByRole('spinbutton', { name: 'Vertical jump size (%)', exact: true }).fill('125');
	await dialog.screenshot({ path: test.info().outputPath('line-jump-page-settings.png') });
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	const layout = () =>
		viewer.evaluate(
			(node) =>
				(
					node as unknown as {
						controller: { state: { document: { pages: { layout: unknown }[] } } };
					}
				).controller.state.document.pages[0]!.layout,
		);
	const expected = {
		lineJumpCode: 1,
		lineJumpStyle: 2,
		lineJumpDirectionX: 2,
		lineJumpDirectionY: 1,
		lineJumpFactorX: 0.5,
		lineJumpFactorY: 1.25,
	};
	await expect.poll(layout).toMatchObject(expected);
	const jumps = viewer.locator('path[data-line-jumps]');
	await expect(jumps).toHaveCount(1);
	const path = await jumps.getAttribute('d');
	expect(path?.match(/M/g)).toHaveLength(2);
	await viewer.locator('.qat [data-command="undo"]').click();
	await expect(jumps).toHaveCount(0);
	await viewer.locator('.qat [data-command="redo"]').click();
	await expect.poll(layout).toMatchObject(expected);
	await expect(jumps).toHaveAttribute('d', path!);
	const printed = await viewer.evaluate(async (node) => {
		const load = (url: string) => import(/* @vite-ignore */ url);
		const { createPrintSnapshot } = await load('/test-api.js');
		const state = (node as unknown as { controller: { state: { document: unknown } } }).controller
			.state;
		return (await createPrintSnapshot(state.document, { pageIndices: [0] })).pages[0].svg;
	});
	expect(printed).toContain(`d="${path}"`);
});

test('connector jump overrides are editable in Format Shape with page jumps disabled', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const line = (id: string, x1: number, y1: number, x2: number, y2: number) =>
		shape(
			id,
			cell('BeginX', x1) +
				cell('BeginY', y1) +
				cell('EndX', x2) +
				cell('EndY', y2) +
				cell('ShapeRouteStyle', 16),
		);
	const bytes = await fixture({
		pages: [
			{
				id: '0',
				pageCells: cell('LineJumpCode', 0),
				contents: `<Shapes>${line('1', 1, 2, 5, 2)}${line('2', 3, 0, 3, 4)}</Shapes>`,
			},
		],
	});
	await page.locator('#file').setInputFiles({
		name: 'jump-overrides.vsdx',
		mimeType: 'application/octet-stream',
		buffer: Buffer.from(bytes),
	});
	await expect(page.locator('#file-name')).toHaveText('jump-overrides.vsdx');
	const viewer = page.locator('visio-viewer');
	await expect(viewer.locator('svg.paper')).toBeVisible();
	await viewer.evaluate((node) => {
		(
			node as unknown as {
				controller: { selectShape(value: { id: string; name: string; pageId: string }): void };
			}
		).controller.selectShape({ id: '1', name: 'Horizontal connector', pageId: '0' });
	});
	await viewer
		.locator(
			'[data-ribbon-tab]:not([data-ribbon-tab="shape-format"]) office-ui-ribbon-group[launcher="shape-styles-dialog"] .launcher',
		)
		.click();
	const pane = viewer.locator('.format-pane');
	await pane.getByRole('combobox', { name: 'Add line jumps', exact: true }).selectOption('2');
	const jumps = viewer.locator('path[data-line-jumps]');
	await expect(jumps).toHaveCount(1);
	await pane.getByRole('combobox', { name: 'Jump style', exact: true }).selectOption('2');
	await pane.screenshot({ path: test.info().outputPath('connector-jump-settings.png') });
	await expect.poll(async () => (await jumps.getAttribute('d'))?.match(/M/g)?.length).toBe(2);
	await viewer.locator('.qat [data-command="undo"]').click();
	await expect(jumps).toHaveAttribute('d', / A /);
	await viewer.locator('.qat [data-command="redo"]').click();
	await expect.poll(async () => (await jumps.getAttribute('d'))?.match(/M/g)?.length).toBe(2);
	await pane.getByRole('combobox', { name: 'Add line jumps', exact: true }).selectOption('4');
	await expect(jumps).toHaveCount(0);
});
