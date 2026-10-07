import { expect, test } from '@playwright/test';
import native from '../../src/core/chart/excel-chart-title-layout.json' with { type: 'json' };
import { loadXlsx, chartView } from 'ooxml-core/xlsx';
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';
import { nativeChartFixture } from './native-chart-fixture';

for (const framework of FRAMEWORKS)
	for (const sample of native.cases)
		test(`native title layout ${sample.referenceName} survives a type edit in ${framework}`, async ({
			page,
		}) => {
			const errors = pageErrors(page);
			await openLanding(page, framework);
			const buffer = await nativeChartFixture(sample.parts, 'Aptos Narrow', 'Aptos Display', {
				width: (sample.layoutGeometry.chart.widthPt * 4) / 3,
				height: (sample.layoutGeometry.chart.heightPt * 4) / 3,
			});
			const original = await loadXlsx(buffer);
			const source = original.sheets[0]!.drawings[0]!;
			if (source.kind !== 'chart') throw new Error('Missing source chart');
			const expected = chartView(original, 0, source, () => []);
			const name = `${sample.referenceName}.xlsx`;
			await page.locator('#landing-file').setInputFiles({
				name,
				mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				buffer,
			});
			const host = editor(page);
			await expect(host.getByText(name, { exact: true })).toBeVisible();
			const verify = async () => {
				const svg = host.locator('.xg-chart svg');
				if (!expected.titleLayout) {
					await expect(svg.locator('g[data-chart-title-layout]')).toHaveCount(0);
					return;
				}
				const group = svg.locator('g[data-chart-title-layout]');
				await expect(group).toHaveCount(1);
				await expect
					.poll(() =>
						group.evaluate((node, reference) => {
							const titleGroup = node as SVGGElement;
							const span = node.querySelector('tspan')!;
							const svg = titleGroup.ownerSVGElement!;
							const matrix = titleGroup.getCTM()!;
							const point = svg.createSVGPoint();
							point.x = Number(span.getAttribute('x'));
							point.y = Number(span.getAttribute('y'));
							const position = point.matrixTransform(matrix);
							return Math.max(
								Math.abs(position.x - ((reference.leftPt + 4) * 4) / 3 - 4),
								Math.abs(position.y - ((reference.topPt + 4) * 4) / 3 - 28),
							);
						}, sample.titleGeometry),
					)
					.toBeLessThan(0.1);
				await expect
					.poll(async () => (await group.textContent())?.replaceAll(' ', ''))
					.toBe(expected.title!.replaceAll(' ', ''));
				await expect
					.poll(() => group.evaluate((node) => node.nextElementSibling === null))
					.toBe(true);
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
			if (chart.kind !== 'chart') throw new Error('Missing saved chart');
			expect(chart.chartType).toBe('line');
			const actual = chartView(saved, 0, chart, () => []);
			expect(actual.titleLayout).toEqual(expected.titleLayout);
			expect(actual.titleOverlay).toBe(expected.titleOverlay);
			expect(actual.title).toBe(expected.title);
			expect(errors).toEqual([]);
		});
