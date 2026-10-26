import { expect, test, type Locator } from '@playwright/test';
import JSZip from 'jszip';
import { createWorkbook, createEditSession, saveXlsx } from 'ooxml-core/xlsx';
import { parseChartSpace } from 'ooxml-core/chart';
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';

async function saved(host: Locator) {
	const bytes = Uint8Array.from(
		await host.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		),
	);
	const xml = await (await JSZip.loadAsync(bytes)).file('xl/charts/chart1.xml')!.async('string');
	return { bytes, xml, space: parseChartSpace(xml).chartSpace };
}

for (const framework of FRAMEWORKS)
	test(`advanced chart formatting, history and fresh reopen in ${framework}`, async ({ page }) => {
		test.setTimeout(120_000);
		await page.setViewportSize({ width: 1600, height: 1000 });
		const errors = pageErrors(page);
		await openLanding(page, framework);
		const book = createWorkbook();
		createEditSession(book).addChart(0, {
			chartType: 'column',
			title: 'Sales',
			showLegend: true,
			anchor: {
				from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 },
				ext: { cx: 600 * 9525, cy: 400 * 9525 },
			},
			series: [{ name: 'Sales', categories: ['A', 'B', 'C'], values: [10, 20, 15] }],
		});
		await page.locator('#landing-file').setInputFiles({
			name: 'advanced.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: Buffer.from(await saveXlsx(book)),
		});
		const host = editor(page);
		const pane = host.locator('office-ui-chart-format-pane');
		const open = async () => {
			await expect(host.locator('.xg-chart')).toBeVisible();
			await host.locator('[data-chart-series="0"][data-chart-point="0"] rect').first().click();
			await host.getByRole('tab', { name: 'Chart Design', exact: true }).click();
			await host.getByRole('button', { name: 'Format Chart Area', exact: true }).click();
			await expect(pane).toBeVisible();
		};
		await open();
		await pane.locator('[data-chart-format-target]').selectOption('series:0');
		await pane.locator('[data-chart-format-field="trendline.type"] select').selectOption('poly');
		const order = pane.locator('[data-chart-format-field="trendline.order"] input');
		await order.fill('3');
		await order.dispatchEvent('change');
		await pane
			.locator('[data-chart-format-field="errorBars.type"] select')
			.selectOption('percentage');
		const value = pane.locator('[data-chart-format-field="errorBars.value"] input');
		await value.fill('5');
		await value.dispatchEvent('change');
		await host.evaluate((node) => (node as unknown as { undo(): void }).undo());
		await expect(value).toHaveValue('1');
		await host.evaluate((node) => (node as unknown as { redo(): void }).redo());
		await expect(value).toHaveValue('5');
		await pane.locator('[data-chart-format-target]').selectOption('chartArea');
		await pane.getByRole('checkbox', { name: 'Outer shadow', exact: true }).check();
		const blur = pane.locator('[data-chart-format-field="shadow.blur"] input[type="number"]');
		await blur.fill('6');
		await blur.dispatchEvent('change');
		await pane.getByRole('combobox', { name: 'Fill', exact: true }).selectOption('pattern');
		await pane
			.locator('[data-chart-format-field="fill.patternPreset"] select')
			.selectOption('cross');
		await pane.locator('[data-chart-format-target]').selectOption('axis:1');
		const title = pane.locator('[data-chart-format-field="axis.title"] input');
		await title.fill('Revenue');
		await title.dispatchEvent('change');
		const result = await saved(host);
		expect(result.space.plotArea.groups[0]!.series[0]!.trendlines?.[0]).toMatchObject({
			trendlineType: 'poly',
			order: 3,
		});
		expect(result.space.plotArea.groups[0]!.series[0]!.errorBars?.[0]).toMatchObject({
			valueType: 'percentage',
			value: 5,
		});
		expect(result.space.spPr?.fill).toMatchObject({ kind: 'pattern', preset: 'cross' });
		expect(result.space.spPr?.effectsXml).toContain('blurRad="76200"');
		expect(result.space.plotArea.axes[1]?.title?.text).toBe('Revenue');
		await openLanding(page, framework);
		await page.locator('#landing-file').setInputFiles({
			name: 'reopened.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: Buffer.from(result.bytes),
		});
		await open();
		await pane.locator('[data-chart-format-target]').selectOption('series:0');
		await expect(pane.locator('[data-chart-format-field="trendline.type"] select')).toHaveValue(
			'poly',
		);
		await expect(pane.locator('[data-chart-format-field="errorBars.value"] input')).toHaveValue(
			'5',
		);
		expect(errors).toEqual([]);
	});
