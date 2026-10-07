import { expect, test } from '@playwright/test';
import { createWorkbook, createEditSession, saveXlsx, loadXlsx } from 'ooxml-core/xlsx';
import { parseXml, NS } from 'ooxml-core/xml';
import {
	parseDrawingFill,
	withDrawingGradientGeometry,
	resolveDrawingColor,
} from 'ooxml-core/diagram';
import { buildChartGradientDef, resolveChartGradient } from 'ooxml-core/chart';
import native from '../../src/core/chart/__fixtures__/native-gradient-series-path-profiles.json' with { type: 'json' };
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';

for (const framework of FRAMEWORKS)
	test(`circle and shape series editing in ${framework}`, async ({ page }) => {
		const errors = pageErrors(page);
		await openLanding(page, framework);
		for (const chartType of ['column', 'bar'] as const) {
			const book = createWorkbook(),
				fill = parseDrawingFill(
					parseXml(`<a:spPr xmlns:a="${NS.a}">${native.cases[0]!.fillXml}</a:spPr>`)
						.documentElement,
				)!;
			createEditSession(book).addChart(0, {
				chartType,
				showLegend: true,
				series: [{ name: 'Sales', categories: ['A', 'B', 'C'], values: [10, 20, 30], fill }],
				anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
			});
			await page.locator('#landing-file').setInputFiles({
				name: `paths-${chartType}.xlsx`,
				mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				buffer: Buffer.from(await saveXlsx(book)),
			});
			const host = editor(page);
			await expect(host.getByText(`paths-${chartType}.xlsx`, { exact: true })).toBeVisible();
			await host.evaluate((node) => ((node as unknown as { readOnly: boolean }).readOnly = false));
			await host.locator('g[data-chart-series="0"][data-chart-point="0"] rect').first().dblclick();
			const pane = host.getByRole('complementary', { name: 'Format Data Series' }),
				type = pane.getByRole('combobox', { name: 'Type', exact: true });
			for (const path of ['circle', 'shape']) {
				await type.selectOption(path);
				await expect(pane.getByRole('spinbutton', { name: 'Angle', exact: true })).toBeDisabled();
				const paints = host.locator(
					`${path === 'circle' ? 'radialGradient' : 'pattern'}[id$="-s0"]`,
				);
				await expect(paints).toHaveCount(4);
				for (const label of [
					'From Center',
					'From Top Left Corner',
					'From Top Right Corner',
					'From Bottom Left Corner',
					'From Bottom Right Corner',
				]) {
					await pane.getByRole('button', { name: 'Direction', exact: true }).click();
					const popup = host.getByRole('dialog', { name: 'Direction', exact: true });
					await expect(popup.getByRole('button')).toHaveCount(5);
					await popup.getByRole('button', { name: label, exact: true }).click();
					await expect(popup).toBeHidden();
					await expect(paints).toHaveCount(4);
				}
				const paintMarkup = () =>
					paints.evaluateAll((nodes) =>
						nodes.map((n) => n.outerHTML.replace(/xlsx-chart-\d+/g, 'chart')),
					);
				if (path === 'circle') {
					await expect(paints.first()).toHaveAttribute('cx', '1');
					await expect(paints.first()).toHaveAttribute('cy', '1');
				} else {
					if (fill.kind !== 'gradient') throw new Error('Expected gradient');
					const expected = buildChartGradientDef(
						'native',
						resolveChartGradient(
							withDrawingGradientGeometry(fill, 'shape', 'bottom-right'),
							(color) => resolveDrawingColor(color),
						),
						{ width: 1, height: 1, shape: 'rect' },
					);
					if (expected.kind !== 'rectPath') throw new Error('Expected shape paint');
					await expect(paints.first().locator('image')).toHaveAttribute('href', expected.href);
				}
				const before = await paintMarkup();
				const stops = pane.getByRole('group', { name: 'Gradient stops', exact: true });
				await stops.scrollIntoViewIfNeeded();
				const strip = await stops.locator('.office-gradient-stop-paint').boundingBox(),
					marker = await stops
						.getByRole('button', { name: 'Gradient stop 1', exact: true })
						.boundingBox();
				if (!strip || !marker) throw new Error('Expected visible stop strip');
				await page.mouse.move(marker.x + marker.width / 2, marker.y + marker.height / 2);
				await page.mouse.down();
				await page.mouse.move(
					marker.x + marker.width / 2 + strip.width * 0.4,
					marker.y + marker.height / 2,
					{ steps: 3 },
				);
				await expect(pane.getByRole('spinbutton', { name: 'Position', exact: true })).toHaveValue(
					'40',
				);
				const during = await paintMarkup();
				for (let i = 0; i < 4; i++) expect(during[i]).not.toBe(before[i]);
				const saved = await host.evaluate(async (node) =>
					Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
				);
				const chart = (await loadXlsx(new Uint8Array(saved))).sheets[0]!.drawings[0]!;
				if (chart.kind !== 'chart' || chart.series[0]!.fill?.kind !== 'gradient')
					throw new Error('Expected gradient');
				expect(chart.series[0]!.fill.path).toBe(path);
				expect(chart.series[0]!.fill.stops[0]!.position).toBe(0);
				expect(chart.series[0]!.fill.fillToRect).toEqual({ l: 1, t: 1, r: 0, b: 0 });
				expect(chart.series[0]!.fill.tileRect).toEqual({ l: 0, t: 0, r: -1, b: -1 });
				await page.keyboard.press('Escape');
				await page.mouse.up();
				expect(await paintMarkup()).toEqual(before);
				const transparency = pane.getByRole('spinbutton', { name: 'Transparency', exact: true });
				await transparency.fill('37');
				await transparency.dispatchEvent('change');
				await expect.poll(paintMarkup).not.toEqual(before);
				await host.evaluate((node) => (node as unknown as { undo(): void }).undo());
				await expect.poll(paintMarkup).toEqual(before);
			}
			await host.evaluate((node) => ((node as unknown as { readOnly: boolean }).readOnly = true));
			await expect(type).toBeDisabled();
			await expect(pane.getByRole('button', { name: 'Direction', exact: true })).toBeDisabled();
		}
		expect(errors).toEqual([]);
	});
