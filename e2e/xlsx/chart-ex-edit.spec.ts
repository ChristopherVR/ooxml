import { readFileSync } from 'node:fs';
import path from 'node:path';
import { expect, test } from '@playwright/test';
import { loadXlsx, saveXlsx, xlsxChartExDataSheet } from 'ooxml-core/xlsx';
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';
async function fixture() {
	const book = await loadXlsx(
		readFileSync(
			path.join(import.meta.dirname, '../../src/core/xlsx/__fixtures__/excel-chart-families.xlsx'),
		),
	);
	const drawing = book.sheets[0]!.drawings.find(
		(d) => d.kind === 'unsupported' && d.chartExSpace?.plotArea.series[0]?.layoutId === 'waterfall',
	)!;
	drawing.anchor = {
		from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 },
		ext: { cx: 600 * 9525, cy: 350 * 9525 },
	};
	book.sheets[0]!.drawings = [drawing];
	return saveXlsx(book);
}
for (const framework of FRAMEWORKS)
	test(`${framework}: ChartEx data growth and options save and reload`, async ({ page }) => {
		test.setTimeout(120000);
		const errors = pageErrors(page);
		await page.setViewportSize({ width: 1500, height: 1000 });
		await openLanding(page, framework);
		await page.locator('#landing-file').setInputFiles({
			name: 'chartex.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: Buffer.from(await fixture()),
		});
		const host = editor(page);
		await expect(host.getByText('chartex.xlsx', { exact: true })).toBeVisible();
		await host.locator('.xg-obj[data-index="0"]').first().dblclick();
		const options = host.locator('office-ui-chart-ex-options');
		await expect(options).toBeVisible();
		await options.getByLabel('Chart title', { exact: true }).fill('Browser ChartEx');
		await options.getByLabel('Chart title', { exact: true }).press('Tab');
		const sheet = host.locator('office-ui-chart-data-sheet');
		const cell = sheet.locator('input[data-row="1"][data-col="1"]');
		await cell.fill('77');
		await cell.press('Enter');
		await sheet.getByRole('button', { name: 'Add category', exact: true }).click();
		const pane = host.locator('office-ui-chart-format-pane');
		await pane.locator('[data-chart-format-target]').selectOption('series:0');
		await pane.getByRole('radio', { name: 'Pattern fill', exact: true }).click();
		await pane
			.locator('[data-chart-format-field="fill.patternPreset"] select')
			.selectOption('cross');
		const bytes = Uint8Array.from(
			await host.evaluate(async (element) =>
				Array.from(await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
			),
		);
		const book = await loadXlsx(bytes);
		expect(book.sheets).toHaveLength(2);
		const data = xlsxChartExDataSheet(book, 0, 0);
		expect(data.series[0]!.values[0]).toBe(77);
		expect(data.categories).toHaveLength(7);
		expect(book.sheets[0]!.drawings[0]!.kind === 'unsupported').toBe(true);
		const drawing = book.sheets[0]!.drawings[0]!;
		expect(
			drawing.kind === 'unsupported' && drawing.chartExSpace!.plotArea.series[0]!.spPr!.fill,
		).toMatchObject({ kind: 'pattern', preset: 'cross' });
		await openLanding(page, framework);
		await page.locator('#landing-file').setInputFiles({
			name: 'chartex-saved.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: Buffer.from(bytes),
		});
		await host.locator('.xg-obj[data-index="0"]').first().dblclick();
		await expect(
			host.locator('office-ui-chart-ex-options').getByLabel('Chart title', { exact: true }),
		).toHaveValue('Browser ChartEx');
		await expect(
			host.locator('office-ui-chart-data-sheet input[data-row="1"][data-col="1"]'),
		).toHaveValue('77');
		expect(errors).toEqual([]);
	});
