import { expect, test } from '@playwright/test';
import { parseVsdx } from 'ooxml-core/visio';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { createVsdxFixture } from './fixture.mjs';
import { openDemo } from './demo-page';

const HEADER = 'Process Step ID,Description,Next Step ID,Connector Label,Shape Type,Function';
const TABLE = [
	HEADER,
	'1,Receive order,2,,Start,Sales',
	'2,Check stock,"3,4","Yes,No",Decision,Warehouse',
	'3,Ship order,5,,Process,Warehouse',
	'4,Back order,5,,Process,Sales',
	'5,Order complete,,,End,Sales',
	'',
].join('\n');

const ribbon = '[data-ribbon-tab]:not([data-contextual]) ';

test('Data Visualizer draws a flowchart from a table and redraws it from the edited table', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.addInitScript((text) => {
		const state = globalThis as unknown as { tableText: string };
		state.tableText = text;
		(window as unknown as { showOpenFilePicker: unknown }).showOpenFilePicker = async () => [
			{ getFile: async () => new File([state.tableText], 'process.csv', { type: 'text/csv' }) },
		];
	}, TABLE);
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'Process.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Process target'),
	});
	await expect(viewer.locator('svg.paper')).toContainText('Process target');
	await viewer.getByRole('tab', { name: 'Data', exact: true }).click();
	await expect(viewer.locator(`${ribbon}[command="visualizer-refresh"]`)).toHaveAttribute(
		'disabled',
		'',
	);

	// Basic flowchart from the table.
	await viewer.getByRole('button', { name: 'Data Visualizer', exact: true }).click();
	await viewer.locator('office-ui-menu-item[command="visualizer-basic"]').click();
	await expect(viewer.locator('[data-status]')).toHaveText(/Drew 5 process steps from process.csv/);
	const paper = viewer.locator('svg.paper');
	await expect(paper).toContainText('Check stock');
	await expect(paper).toContainText('Order complete');
	await expect(paper).toContainText('Yes');
	await page.screenshot({ path: test.info().outputPath('data-visualizer-basic.png') });
	const saved = async () => {
		const bytes = await viewer.evaluate((node) =>
			Array.from((node as VisioViewerElement).exportVsdx().bytes),
		);
		return (await parseVsdx(new Uint8Array(bytes))).pages[0]!;
	};
	const stepIds = (drawing: Awaited<ReturnType<typeof saved>>) =>
		new Map(
			drawing.shapes.flatMap((shape) => {
				const row = shape.shapeData?.find((item) => item.name === '_ooxmlVisualizer');
				return row?.label === 'step' ? [[String(row.value), shape.id] as const] : [];
			}),
		);
	const first = await saved();
	const firstIds = stepIds(first);
	// Connector labels are the connectors' own text, not separate text boxes.
	const labelled = first.shapes.filter((shape) => shape.text.plainText === 'Yes');
	expect(labelled).toHaveLength(1);
	expect(labelled[0]!.kind).toBe('connector');

	// Edit the table: rename a step, add one, and redraw.
	await page.evaluate(() => {
		const state = globalThis as unknown as { tableText: string };
		state.tableText = state.tableText
			.replace('Ship order', 'Dispatch order')
			.replace(
				'5,Order complete,,,End,Sales',
				'5,Order complete,6,,End,Sales\n6,Archive,,,Process,Sales',
			);
	});
	await viewer.getByRole('button', { name: 'Data Visualizer', exact: true }).click();
	await viewer.locator('office-ui-menu-item[command="visualizer-refresh"]').click();
	await expect(viewer.locator('[data-status]')).toHaveText(/Redrew 6 process steps/);
	await expect(paper).toContainText('Dispatch order');
	await expect(paper).toContainText('Archive');
	await expect(paper).not.toContainText('Ship order');
	const after = await saved();
	// Rows that stayed keep their shapes (same IDs); only the new row adds one.
	const afterIds = stepIds(after);
	for (const key of ['1', '2', '3', '4', '5']) expect(afterIds.get(key)).toBe(firstIds.get(key));
	const marked = after.shapes.filter((shape) =>
		shape.shapeData?.some((row) => row.name === '_ooxmlVisualizer' && row.label === 'step'),
	);
	expect(marked).toHaveLength(6);
	expect(marked[0]!.shapeData!.find((row) => row.name === 'ProcessStepID')?.value).toBeDefined();
	await page.screenshot({ path: test.info().outputPath('data-visualizer-refreshed.png') });
	// The whole redraw is one undo step.
	await viewer.locator('.qat [data-command="undo"]').click();
	await expect(paper).toContainText('Ship order');
	await expect(paper).not.toContainText('Archive');
	expect(errors).toEqual([]);
});

test('Data Visualizer lays steps out in swimlanes by Function', async ({ page }) => {
	await page.addInitScript((text) => {
		(window as unknown as { showOpenFilePicker: unknown }).showOpenFilePicker = async () => [
			{ getFile: async () => new File([text], 'process.csv', { type: 'text/csv' }) },
		];
	}, TABLE);
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'Lanes.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Lanes target'),
	});
	await expect(viewer.locator('svg.paper')).toContainText('Lanes target');
	await viewer.getByRole('tab', { name: 'Data', exact: true }).click();
	await viewer.getByRole('button', { name: 'Data Visualizer', exact: true }).click();
	await viewer.locator('office-ui-menu-item[command="visualizer-cross-functional"]').click();
	await expect(viewer.locator('[data-status]')).toHaveText(/Drew 5 process steps/);
	const paper = viewer.locator('svg.paper');
	await expect(paper).toContainText('Sales');
	await expect(paper).toContainText('Warehouse');
	await expect(paper).toContainText('Check stock');
	await page.screenshot({ path: test.info().outputPath('data-visualizer-lanes.png') });
});

test('Data Graphic Fields and Advanced Data Graphics edit the items of a shape', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.addInitScript((text) => {
		(window as unknown as { showOpenFilePicker: unknown }).showOpenFilePicker = async () => [
			{ getFile: async () => new File([text], 'process.csv', { type: 'text/csv' }) },
		];
	}, TABLE);
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'Graphics.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Graphics target'),
	});
	await expect(viewer.locator('svg.paper')).toContainText('Graphics target');
	await viewer.getByRole('tab', { name: 'Data', exact: true }).click();
	await viewer.getByRole('button', { name: 'Data Visualizer', exact: true }).click();
	await viewer.locator('office-ui-menu-item[command="visualizer-basic"]').click();
	await expect(viewer.locator('[data-status]')).toHaveText(/Drew 5 process steps/);
	// The fixture's own shape is 1; the steps follow it, so "Check stock" is shape 4.
	const step = viewer.locator('svg.paper [data-shape-id="4"]');
	await step.click();
	const saved = async () => {
		const bytes = await viewer.evaluate((node) =>
			Array.from((node as VisioViewerElement).exportVsdx().bytes),
		);
		return (await parseVsdx(new Uint8Array(bytes))).pages[0]!;
	};
	const graphics = async () =>
		(await saved()).shapes
			.flatMap((shape) => shape.shapeData?.filter((row) => row.name === '_ooxmlDataGraphic') ?? [])
			.map((row) => `${row.label}:${row.prompt}`)
			.sort();

	// Data Graphic Fields: the pane lists the selected shape's fields; a check shows it as text.
	await viewer.locator(`${ribbon}[data-check="data-graphic-fields"]`).click();
	const pane = viewer.locator('.graphic-fields');
	await expect(pane).toBeVisible();
	await expect(pane.locator('label')).toContainText([/Process Step ID/, /Description/]);
	await pane.locator('input[data-field="Description"]').check();
	await expect(viewer.locator('[data-status]')).toHaveText(/Showing Description/);
	await expect.poll(graphics).toEqual(['text:Description']);
	await pane.locator('input[data-field="Function"]').check();
	await expect.poll(graphics).toEqual(['text:Description', 'text:Function']);
	await pane.locator('input[data-field="Function"]').uncheck();
	await expect.poll(graphics).toEqual(['text:Description']);

	// Advanced Data Graphics: list, add, edit and delete items.
	await viewer.locator(`${ribbon}[command="advanced-data-graphics"] button`).click();
	const dialog = viewer.locator('.data-graphic-items-dialog');
	const list = dialog.locator('select[name="items"] option');
	await expect(list).toHaveText(['Text Callout: Description']);
	await dialog.locator('[label="New Item"] button').click();
	const form = viewer.locator('.data-graphics-dialog');
	await form.locator('[name="field"]').selectOption('ProcessStepID');
	await form.locator('[name="kind"]').selectOption('bar');
	await form.locator('[name="position"]').selectOption('top');
	await form.locator('[label="Apply"] button').click();
	await expect(viewer.locator('[data-status]')).toHaveText(/Added the item/);
	await expect(list).toHaveText(['Text Callout: Description', 'Data Bar: Process Step ID (top)']);
	await expect
		.poll(graphics)
		.toEqual(['bar:ProcessStepID', 'bar:ProcessStepID', 'text:Description']);
	await page.screenshot({ path: test.info().outputPath('advanced-data-graphics.png') });
	await dialog.locator('select[name="items"]').selectOption('1');
	await dialog.locator('[label="Edit Item"] button').click();
	await form.locator('[name="kind"]').selectOption('icon');
	await form.locator('[label="Apply"] button').click();
	await expect(viewer.locator('[data-status]')).toHaveText(/Changed the data graphic item/);
	await expect(list).toHaveText(['Text Callout: Description', 'Icon Set: Process Step ID (top)']);
	await dialog.locator('select[name="items"]').selectOption('1');
	await dialog.locator('[label="Delete Item"] button').click();
	await expect(list).toHaveText(['Text Callout: Description']);
	await dialog.locator('[label="Close"] button').click();
	await expect.poll(graphics).toEqual(['text:Description']);
	expect(errors).toEqual([]);
});
