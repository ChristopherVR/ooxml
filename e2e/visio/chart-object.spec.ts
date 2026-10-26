import { expect, test, type Locator } from '@playwright/test';
import { chartSpaceDataSheet } from 'ooxml-core/chart';
import { readEmbeddedChartSpace } from 'ooxml-core/xlsx';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { createVsdxFixture } from './fixture.mjs';
import { openDemo } from './demo-page';

/** The embedded workbook of the page's chart object, read in the test process. */
async function embeddedWorkbook(viewer: Locator): Promise<Uint8Array> {
	const bytes = await viewer.evaluate((node) => {
		const state = (node as VisioViewerElement).controller.state;
		const shapes = state.document?.pages[state.pageIndex]?.shapes ?? [];
		const chart = shapes.find((shape) => shape.embeddedObject);
		return chart ? Array.from(chart.embeddedObject!.bytes) : [];
	});
	return Uint8Array.from(bytes);
}

test('Insert > Chart, then Edit Data and Format Chart re-embed the chart workbook', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await expect(viewer.locator('svg.paper')).toContainText('Release workflow');
	await page.locator('#file').setInputFiles({
		name: 'Chart.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Chart target'),
	});
	await expect(page.locator('#file-name')).toHaveText('Chart.vsdx');
	await expect(viewer.locator('svg.paper')).toContainText('Chart target');
	await viewer.getByRole('tab', { name: 'Insert', exact: true }).click();
	await viewer.locator('[command="chart"] button').click();
	await viewer.locator('[command="chart-dialog-insert"] button').click();
	await expect(viewer.locator('[data-status]')).toHaveText(/Inserted chart/);
	const chart = viewer.locator('svg.paper [data-shape-id]:has([data-embedded-chart])').first();
	await expect(chart).toBeVisible();

	// Edit Data: the shared data sheet, one value changed.
	await chart.click({ button: 'right' });
	await viewer.locator('[command="ctx-chart-data"]').click();
	const dialog = viewer.locator('.chart-edit-dialog');
	await expect(dialog).toHaveAttribute('heading', 'Edit Data');
	const cell = dialog.locator('office-ui-chart-data-sheet input[data-row="1"][data-col="1"]');
	await cell.fill('9');
	await cell.press('Enter');
	await expect(dialog.locator('.chart-edit-preview svg')).toBeVisible();
	await dialog.locator('[command="chart-edit-ok"] button').click();
	await expect(viewer.locator('[data-status]')).toHaveText(/Updated the chart/);
	let found = (await readEmbeddedChartSpace(await embeddedWorkbook(viewer)))!;
	expect(chartSpaceDataSheet(found.space).series[0]?.values[0]).toBe(9);

	// Format Chart: the shared Format pane, the first series filled red.
	await chart.click({ button: 'right' });
	await viewer.locator('[command="ctx-chart-format"]').click();
	await expect(dialog).toHaveAttribute('heading', 'Format Chart');
	const pane = dialog.locator('office-ui-chart-format-pane');
	await pane.locator('[data-chart-format-target]').selectOption('series:0');
	await pane.getByRole('radio', { name: 'Solid fill', exact: true }).click();
	await pane.locator('[data-chart-format-field="fill.color"] input').fill('#ff0000');
	await pane.locator('[data-chart-format-tab="options"]').click();
	await pane.locator('[data-chart-format-field="trendline.type"] select').selectOption('linear');
	await pane.locator('[data-chart-format-field="errorBars.type"] select').selectOption('fixedVal');
	const errorValue = pane.locator('[data-chart-format-field="errorBars.value"] input');
	await errorValue.fill('2');
	await errorValue.dispatchEvent('change');
	await dialog.locator('[command="chart-edit-ok"] button').click();
	// The status still reads "Updated the chart." from Edit Data: wait for the new workbook.
	await expect(dialog).toBeHidden();
	await expect
		.poll(async () => {
			found = (await readEmbeddedChartSpace(await embeddedWorkbook(viewer)))!;
			return found.space.plotArea.groups[0]?.series[0]?.spPr?.fill?.kind === 'solid'
				? found.space.plotArea.groups[0]?.series[0]?.spPr?.fill.color.value
				: undefined;
		})
		.toBe('FF0000');
	expect(found.space.plotArea.groups[0]?.series[0]?.spPr?.fill).toMatchObject({
		kind: 'solid',
		color: { value: 'FF0000' },
	});
	expect(chartSpaceDataSheet(found.space).series[0]?.values[0]).toBe(9);
	expect(found.space.plotArea.groups[0]?.series[0]?.trendlines?.[0]?.trendlineType).toBe('linear');
	expect(found.space.plotArea.groups[0]?.series[0]?.errorBars?.[0]?.value).toBe(2);
	expect(errors).toEqual([]);
});
