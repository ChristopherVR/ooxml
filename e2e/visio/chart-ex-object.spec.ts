import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';
import { chartExSpaceDataSheet } from 'ooxml-core/chart';
import { loadXlsx, saveXlsx, readEmbeddedChartExSpace } from 'ooxml-core/xlsx';
import { createVsdx, editVsdx } from 'ooxml-core/visio';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';
async function fixture() {
	const book = await loadXlsx(
		readFileSync(
			new URL('../../src/core/xlsx/__fixtures__/excel-chart-families.xlsx', import.meta.url),
		),
	);
	for (const sheet of book.sheets)
		sheet.drawings = sheet.drawings.filter(
			(d) =>
				d.kind === 'unsupported' && d.chartExSpace?.plotArea.series[0]?.layoutId === 'waterfall',
		);
	const workbook = await saveXlsx(book);
	return (
		await editVsdx(await createVsdx(), [
			{
				type: 'insert-chart-object',
				pageId: '0',
				shapeId: '1',
				x: 4,
				y: 5,
				width: 6,
				height: 4,
				workbook,
			},
		])
	).bytes;
}
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: embedded ChartEx paints, edits, and preserves its workbook`, async ({
		page,
	}, testInfo) => {
		const errors: string[] = [];
		page.on('pageerror', (e) => errors.push(e.message));
		await page.setViewportSize({ width: 1500, height: 1000 });
		await openDemo(
			page,
			framework === 'vanilla' ? '/demo/?sample=1' : `/demo-${framework}/?sample=1`,
		);
		await page.locator('#file').setInputFiles({
			name: 'ChartEx.vsdx',
			mimeType: 'application/vnd.ms-visio.drawing',
			buffer: Buffer.from(await fixture()),
		});
		const viewer = page.locator('visio-viewer');
		const chart = viewer.locator('svg.paper [data-shape-id]:has([data-embedded-chart])').first();
		await expect(chart.locator('svg [data-chart-point]')).not.toHaveCount(0);
		await chart.click({ button: 'right' });
		await viewer.locator('[command="ctx-chart-data"]').click();
		const dialog = viewer.locator('.chart-edit-dialog');
		await expect(dialog).toHaveAttribute('heading', 'Edit Data');
		const cell = dialog.locator('office-ui-chart-data-sheet input[data-row="1"][data-col="1"]');
		await cell.fill('42');
		await cell.press('Enter');
		await expect(dialog.locator('.chart-edit-preview svg')).toBeVisible();
		await dialog.locator('[command="chart-edit-ok"] button').click();
		await expect(dialog).toBeHidden();
		const data = await viewer.evaluate((node) =>
			Array.from(
				(node as VisioViewerElement).controller.state.document!.pages[0]!.shapes.find(
					(s) => s.embeddedObject,
				)!.embeddedObject!.bytes,
			),
		);
		const found = (await readEmbeddedChartExSpace(Uint8Array.from(data)))!;
		expect(chartExSpaceDataSheet(found.space).series[0]!.values[0]).toBe(42);
		await chart.click({ button: 'right' });
		await viewer.locator('[command="ctx-chart-format"]').click();
		const pane = dialog.locator('office-ui-chart-format-pane');
		await pane.locator('[data-chart-format-target]').selectOption('series:0');
		await pane.getByRole('radio', { name: 'Solid fill', exact: true }).click();
		await pane.locator('[data-chart-format-field="fill.color"] input').fill('#ff0000');
		await dialog.locator('[command="chart-edit-ok"] button').click();
		await expect(dialog).toBeHidden();
		const formatted = await viewer.evaluate((node) =>
			Array.from(
				(node as VisioViewerElement).controller.state.document!.pages[0]!.shapes.find(
					(s) => s.embeddedObject,
				)!.embeddedObject!.bytes,
			),
		);
		expect(
			(await readEmbeddedChartExSpace(Uint8Array.from(formatted)))!.space.plotArea.series[0]!.spPr
				?.fill,
		).toMatchObject({ kind: 'solid', color: { value: 'FF0000' } });
		await testInfo.attach('embedded-chartex.png', {
			body: await chart.screenshot(),
			contentType: 'image/png',
		});
		expect(errors).toEqual([]);
	});
