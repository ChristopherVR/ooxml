import { expect, test } from '@playwright/test';
import native from '../../src/core/chart/excel-chart-title-wrap.json' with { type: 'json' };
import { loadXlsx, chartView } from 'ooxml-core/xlsx';
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';
import { nativeChartFixture } from './native-chart-fixture';

for (const framework of FRAMEWORKS)
	for (const sample of native.cases)
		test(`native wrapped title ${sample.referenceName} survives a type edit in ${framework}`, async ({
			page,
		}) => {
			const errors = pageErrors(page);
			await openLanding(page, framework);
			const buffer = await nativeChartFixture(sample.parts, 'Aptos Narrow', 'Aptos Display', {
				width: (sample.chartGeometry.widthPt * 4) / 3,
				height: (sample.chartGeometry.heightPt * 4) / 3,
			});
			const original = await loadXlsx(buffer);
			const source = original.sheets[0]!.drawings[0]!;
			if (source.kind !== 'chart') throw new Error('Missing source chart');
			const expected = chartView(original, 0, source, () => []).titleText;
			const name = `${sample.referenceName}.xlsx`;
			await page.locator('#landing-file').setInputFiles({
				name,
				mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				buffer,
			});
			const host = editor(page);
			await expect(host.getByText(name, { exact: true })).toBeVisible();
			const verify = async () => {
				const spans = host.locator('.xg-chart svg tspan[data-chart-title-run]');
				await expect(spans.first()).toHaveAttribute('fill', '#FF0000');
				const expectedRows =
					sample.referenceName === 'wrap-short-480'
						? 1
						: sample.referenceName.includes('long-200')
							? 5
							: /long-320|mixed|word-160|word-200/.test(sample.referenceName)
								? 3
								: 2;
				await expect
					.poll(() =>
						spans.evaluateAll((nodes) => new Set(nodes.map((node) => node.getAttribute('y'))).size),
					)
					.toBe(expectedRows);
				const rows = await spans.evaluateAll((nodes) => {
					const lines = new Map<string, string>();
					for (const node of nodes) {
						const y = node.getAttribute('y')!;
						lines.set(y, (lines.get(y) ?? '') + node.textContent);
					}
					return [...lines.values()];
				});
				expect(rows.join('').replaceAll(' ', '')).toBe(source.title!.replaceAll(' ', ''));
				if (sample.referenceName === 'wrap-word-200')
					expect(rows).toEqual(['InternationalBusin', 'essRevenueForec', 'ast']);
				if (sample.referenceName === 'wrap-mixed-320')
					expect(rows).toEqual([
						'Revenue growth and',
						'forecast for the entire',
						'international business region',
					]);
			};
			await verify();
			await host
				.locator('rect[data-chart-part="chartArea"]')
				.first()
				.click({ position: { x: 20, y: 20 } });
			if (sample.referenceName === 'wrap-word-200') {
				const grip = await host.locator('.xg-grip[data-grip="se"]').boundingBox();
				if (!grip) throw new Error('Missing resize grip');
				await page.mouse.move(grip.x + grip.width / 2, grip.y + grip.height / 2);
				await page.mouse.down();
				await page.mouse.move(grip.x + grip.width / 2 + 160, grip.y + grip.height / 2, {
					steps: 5,
				});
				await page.mouse.up();
				const spans = host.locator('.xg-chart svg tspan[data-chart-title-run]');
				await expect(spans).toHaveText(['InternationalBusinessRevenue', 'Forecast']);
				await page.keyboard.press('Control+Z');
				await verify();
				await page.keyboard.press('Control+Y');
				await expect(spans).toHaveText(['InternationalBusinessRevenue', 'Forecast']);
				await page.keyboard.press('Control+Z');
				await verify();
			}
			await host.getByRole('tab', { name: 'Chart Design', exact: true }).click();
			await host.getByRole('button', { name: 'Change Chart Type', exact: true }).click();
			await page.getByRole('button', { name: 'Line', exact: true }).click();
			await page.getByRole('button', { name: 'OK', exact: true }).click();
			await verify();
			const bytes = await host.evaluate(async (node) =>
				Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
			);
			const saved = await loadXlsx(Uint8Array.from(bytes));
			const chart = saved.sheets[0]!.drawings[0]!;
			if (chart.kind !== 'chart') throw new Error('Missing saved chart');
			expect(chart.chartType).toBe('line');
			expect(chart.title).toBe(source.title);
			expect(chartView(saved, 0, chart, () => []).titleText).toEqual(expected);
			expect(errors).toEqual([]);
		});
