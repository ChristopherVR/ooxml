import { expect, test } from '@playwright/test';
import { createWorkbook, createEditSession, saveXlsx, loadXlsx } from 'ooxml-core/xlsx';
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';

for (const framework of FRAMEWORKS)
	test(`axis visibility survives UI export in ${framework}`, async ({ page }) => {
		const errors = pageErrors(page);
		await openLanding(page, framework);
		for (const [axisVisible, labelsVisible] of [
			[true, true],
			[true, false],
			[false, false],
		] as const) {
			const book = createWorkbook();
			createEditSession(book).addChart(0, {
				chartType: 'column',
				showLegend: false,
				anchor: {
					from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 },
					ext: { cx: 600 * 9525, cy: 400 * 9525 },
				},
				series: [{ categories: ['Axis A', 'Axis B'], values: [10, 20] }],
				formatting: {
					sourceXml: '',
					entries: {
						categoryAxis: { sourceXml: '', axisVisible, labelsVisible },
						valueAxis: { sourceXml: '', axisVisible, labelsVisible },
					},
				},
			});
			const name = `axes-${axisVisible}-${labelsVisible}.xlsx`;
			await page.locator('#landing-file').setInputFiles({
				name,
				mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				buffer: Buffer.from(await saveXlsx(book)),
			});
			const host = editor(page);
			await expect(host.getByText(name, { exact: true })).toBeVisible();
			const labels = host.locator('svg text').filter({ hasText: /^Axis [AB]$/ });
			await expect(labels).toHaveCount(labelsVisible ? 2 : 0);
			const bytes = await host.evaluate(async (node) =>
				Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
			);
			const back = await loadXlsx(Uint8Array.from(bytes));
			const chart = back.sheets[0]!.drawings.find((item) => item.kind === 'chart');
			expect(chart?.kind).toBe('chart');
			if (chart?.kind !== 'chart') throw new Error('Missing exported chart');
			for (const part of ['categoryAxis', 'valueAxis'] as const)
				expect(chart.formatting?.entries[part]).toMatchObject({ axisVisible, labelsVisible });
		}
		expect(errors).toEqual([]);
	});
