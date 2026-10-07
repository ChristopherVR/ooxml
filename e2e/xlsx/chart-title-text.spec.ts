import { expect, test } from '@playwright/test';
import native from '../../src/core/chart/excel-chart-title-text.json' with { type: 'json' };
import { loadXlsx, chartView } from 'ooxml-core/xlsx';
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';
import { nativeChartFixture } from './native-chart-fixture';

for (const framework of FRAMEWORKS)
	for (const sample of native.cases)
		test(`native rich title ${sample.referenceName} survives a type edit in ${framework}`, async ({
			page,
		}) => {
			const errors = pageErrors(page);
			await openLanding(page, framework);
			const name = `${sample.referenceName}.xlsx`;
			await page.locator('#landing-file').setInputFiles({
				name,
				mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				buffer: await nativeChartFixture(sample.parts, 'Aptos Narrow'),
			});
			const host = editor(page);
			await expect(host.getByText(name, { exact: true })).toBeVisible();
			const spans = host.locator('.xg-chart svg tspan[data-chart-title-run]');
			const revenue = spans.filter({ hasText: /^Revenue$/ });
			const growth = spans.filter({ hasText: /^growth$/ });
			const verify = async () => {
				await expect(revenue).toHaveAttribute('font-size', '32');
				await expect(revenue).toHaveAttribute('font-weight', 'bold');
				await expect(revenue).toHaveAttribute('fill', '#FF0000');
				await expect(revenue).toHaveAttribute('font-family', /Arial/);
				await expect(growth).toHaveAttribute('font-size', '16');
				await expect(growth).toHaveAttribute('font-style', 'italic');
				await expect(growth).toHaveAttribute('fill', '#0000FF');
				// Newly replaced SVG text can report zero advances before browser layout.
				await expect
					.poll(() =>
						spans.evaluateAll((nodes) => {
							const advances = nodes.slice(0, 3).map((node) => ({
								x: Number(node.getAttribute('x')),
								width: (node as SVGTextContentElement).getComputedTextLength(),
							}));
							const svgWidth = Number(nodes[0]!.closest('svg')!.getAttribute('width'));
							return Math.max(
								...advances
									.slice(1)
									.map((item, index) =>
										Math.abs(item.x - advances[index]!.x - advances[index]!.width),
									),
								Math.abs(
									advances[0]!.x +
										advances.reduce((sum, item) => sum + item.width, 0) / 2 -
										svgWidth / 2,
								),
							);
						}),
					)
					.toBeLessThan(0.5);
				if (sample.referenceName !== 'mixed-single') {
					const forecast = spans.filter({ hasText: /^Forecast$/ });
					await expect(forecast).toHaveAttribute('text-decoration', 'underline');
					await expect(forecast).toHaveAttribute('font-family', /Cambria/);
					expect(Number(await forecast.getAttribute('y'))).toBeGreaterThan(
						Number(await revenue.getAttribute('y')),
					);
				}
			};
			await verify();
			await host
				.locator('rect[data-chart-part="chartArea"]')
				.first()
				.click({ position: { x: 20, y: 20 } });
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
			if (chart.kind !== 'chart') throw new Error('Missing chart');
			expect(chart.chartType).toBe('line');
			const body = chartView(saved, 0, chart, () => []).titleText!;
			expect(
				body.paragraphs.flatMap((paragraph) => paragraph.runs).find((run) => run.text === 'growth')!
					.appearance,
			).toMatchObject({ fontSize: 12, italic: true, color: '#0000FF' });
			if (sample.referenceName === 'mixed-paragraphs') expect(body.paragraphs).toHaveLength(2);
			expect(errors).toEqual([]);
		});
