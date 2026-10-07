import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { createWorkbook, createEditSession, loadXlsx, saveXlsx } from 'ooxml-core/xlsx';
import { OFFICE_GRADIENT_PRESETS } from 'ooxml-core/diagram';
import native from '../../src/core/xlsx/__fixtures__/excel-gradient-presets.json' with { type: 'json' };
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';

for (const framework of FRAMEWORKS)
	test(`native named gradient presets in ${framework}`, async ({ page }) => {
		const errors = pageErrors(page);
		const book = createWorkbook();
		createEditSession(book).addChart(0, {
			chartType: 'column',
			series: [],
			showLegend: false,
			anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		});
		const zip = await JSZip.loadAsync(await saveXlsx(book));
		zip.file('xl/charts/chart1.xml', native.cases[0]!.chartXml);
		await openLanding(page, framework);
		await page.locator('#landing-file').setInputFiles({
			name: 'native-presets.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: await zip.generateAsync({ type: 'nodebuffer' }),
		});
		await editor(page)
			.locator('g[data-chart-series="0"][data-chart-point="0"] rect')
			.first()
			.dblclick();
		const pane = editor(page).getByRole('complementary', { name: 'Format Data Series' });
		const trigger = pane.getByRole('button', { name: 'Preset gradients', exact: true });
		for (const sample of native.cases) {
			await trigger.click();
			const popup = editor(page).getByRole('dialog', { name: 'Preset gradients', exact: true });
			await expect(popup.getByRole('button')).toHaveCount(24);
			await popup
				.getByRole('button', { name: OFFICE_GRADIENT_PRESETS[sample.id - 1]!.label, exact: true })
				.click();
			await expect(popup).toBeHidden();
			expect(await pane.evaluate((node) => node.scrollWidth <= node.clientWidth)).toBe(true);
			await expect(
				pane.getByRole('group', { name: 'Gradient stops' }).getByRole('button'),
			).toHaveCount(sample.stops.length);
			await expect(pane.getByRole('spinbutton', { name: 'Angle', exact: true })).toHaveValue(
				String(sample.angle),
			);
			const bytes = await editor(page).evaluate(async (node) =>
				Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
			);
			const drawing = (await loadXlsx(new Uint8Array(bytes))).sheets[0]!.drawings[0]!;
			if (drawing.kind !== 'chart' || drawing.series[0]!.fill?.kind !== 'gradient')
				throw new Error('Expected gradient');
			const fill = drawing.series[0]!.fill;
			fill.stops.forEach((stop, index) => {
				const expected = sample.stops[index]!;
				expect(stop.position).toBeCloseTo(expected.position, 3);
				expect(stop.color.kind).toBe('srgb');
				const rgb = expected.rgb;
				expect(stop.color.value).toBe(
					[rgb & 255, (rgb >>> 8) & 255, (rgb >>> 16) & 255]
						.map((value) => value.toString(16).padStart(2, '0'))
						.join('')
						.toUpperCase(),
				);
				expect(stop.color.transforms).toEqual([]);
			});
		}
		await editor(page).evaluate((node) => (node as unknown as { undo(): void }).undo());
		await expect(
			pane.getByRole('group', { name: 'Gradient stops' }).getByRole('button'),
		).toHaveCount(native.cases[22]!.stops.length);
		await trigger.focus();
		await trigger.press('ArrowDown');
		const popup = editor(page).getByRole('dialog', { name: 'Preset gradients', exact: true });
		await page.keyboard.press('End');
		await page.keyboard.press('Enter');
		await expect(popup).toBeHidden();
		await expect(
			pane.getByRole('group', { name: 'Gradient stops' }).getByRole('button'),
		).toHaveCount(native.cases[23]!.stops.length);
		await editor(page).evaluate((node) => {
			(node as unknown as { readOnly: boolean }).readOnly = true;
		});
		await expect(trigger).toBeDisabled();
		expect(errors).toEqual([]);
	});
