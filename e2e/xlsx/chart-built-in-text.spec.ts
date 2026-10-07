import { expect, test } from '@playwright/test';
import native from '../../src/core/chart/excel-built-in-text-styles.json' with { type: 'json' };
import { loadXlsx, chartView } from 'ooxml-core/xlsx';
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';
import { nativeChartFixture } from './native-chart-fixture';

for (const framework of FRAMEWORKS)
	for (const id of [2, 102, 141])
		test(`built-in text style ${id} survives a type edit in ${framework}`, async ({ page }) => {
			const errors = pageErrors(page);
			const sample = native.cases.find((item) => item.requestedStyle === id)!;
			await openLanding(page, framework);
			const name = `built-in-${id}.xlsx`;
			await page.locator('#landing-file').setInputFiles({
				name,
				mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				buffer: await nativeChartFixture(sample.parts, sample.titleText.name),
			});
			const host = editor(page);
			await expect(host.getByText(name, { exact: true })).toBeVisible();
			const title = host.locator('.xg-chart svg text').filter({ hasText: 'Native style' });
			await expect(title).toHaveAttribute('font-size', String((sample.titleText.size! * 4) / 3));
			await expect(title).toHaveAttribute('font-weight', 'bold');
			await expect(title).toHaveAttribute('font-family', new RegExp(sample.titleText.name));
			await expect(title).toHaveAttribute('fill', sample.titleText.color);
			await host
				.locator('rect[data-chart-part="chartArea"]')
				.first()
				.click({ position: { x: 20, y: 20 } });
			await host.getByRole('tab', { name: 'Chart Design', exact: true }).click();
			await host.getByRole('button', { name: 'Change Chart Type', exact: true }).click();
			// Dialog content is slotted into the shared control, outside its internal role wrapper.
			await page.getByRole('button', { name: 'Line', exact: true }).click();
			await page.getByRole('button', { name: 'OK', exact: true }).click();
			await expect(title).toHaveAttribute('font-size', String((sample.titleText.size! * 4) / 3));
			const bytes = await host.evaluate(async (n) =>
				Array.from(await (n as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
			);
			const saved = await loadXlsx(Uint8Array.from(bytes));
			const chart = saved.sheets[0]!.drawings[0]!;
			if (chart.kind !== 'chart') throw new Error('Missing chart');
			expect(chart.chartType).toBe('line');
			expect(chart.formatting?.builtInStyle).toBe(id);
			expect(chartView(saved, 0, chart, () => []).appearance?.title).toMatchObject({
				fontSize: sample.titleText.size,
				bold: true,
				typeface: sample.titleText.name,
				color: sample.titleText.color,
			});
			expect(errors).toEqual([]);
		});
