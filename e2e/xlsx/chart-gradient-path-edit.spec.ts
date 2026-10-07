import { expect, test } from '@playwright/test';
import { createWorkbook, createEditSession, saveXlsx, loadXlsx } from 'ooxml-core/xlsx';
import { parseXml, NS } from 'ooxml-core/xml';
import { parseDrawingFill } from 'ooxml-core/diagram';
import native from '../../src/core/chart/__fixtures__/native-gradient-path-profiles.json' with { type: 'json' };
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';

for (const framework of FRAMEWORKS)
	test(`rectangular chart gradient live editing in ${framework}`, async ({ page }) => {
		const errors = pageErrors(page);
		const book = createWorkbook();
		const fill = parseDrawingFill(
			parseXml(`<a:spPr xmlns:a="${NS.a}">${native.cases[0]!.fillXml}</a:spPr>`).documentElement,
		)!;
		createEditSession(book).addChart(0, {
			chartType: 'column',
			showLegend: false,
			series: [{ name: 'Sales', categories: ['A', 'B'], values: [10, 20], fill }],
			anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		});
		await openLanding(page, framework);
		await page.locator('#landing-file').setInputFiles({
			name: 'rect-path.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: Buffer.from(await saveXlsx(book)),
		});
		const host = editor(page);
		await host.locator('g[data-chart-series="0"][data-chart-point="0"] rect').first().dblclick();
		const pane = host.getByRole('complementary', { name: 'Format Data Series' });
		await expect(pane.getByRole('spinbutton', { name: 'Angle', exact: true })).toBeDisabled();
		await expect(pane.getByRole('button', { name: 'Direction', exact: true })).toBeEnabled();
		const paint = host.locator('pattern[id$="-s0"] image').first();
		const original = await paint.getAttribute('href');
		const stops = pane.getByRole('group', { name: 'Gradient stops', exact: true });
		await stops.scrollIntoViewIfNeeded();
		const track = await stops.locator('.office-gradient-stop-paint').boundingBox();
		const marker = await stops
			.getByRole('button', { name: 'Gradient stop 1', exact: true })
			.boundingBox();
		if (!track || !marker) throw new Error('Expected visible stop strip');
		const x = marker.x + marker.width / 2,
			y = marker.y + marker.height / 2;
		await page.mouse.move(x, y);
		await page.mouse.down();
		await page.mouse.move(x + track.width * 0.4, y, { steps: 3 });
		await expect(pane.getByRole('spinbutton', { name: 'Position', exact: true })).toHaveValue('40');
		await expect(paint).not.toHaveAttribute('href', original!);
		const saved = await host.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		const preview = (await loadXlsx(new Uint8Array(saved))).sheets[0]!.drawings[0]!;
		if (preview.kind !== 'chart' || preview.series[0]!.fill?.kind !== 'gradient')
			throw new Error('Expected chart gradient');
		expect(preview.series[0]!.fill.stops[0]!.position).toBe(0);
		await page.keyboard.press('Escape');
		await page.mouse.up();
		await expect(paint).toHaveAttribute('href', original!);
		const transparency = pane.getByRole('spinbutton', { name: 'Transparency', exact: true });
		await transparency.fill('37');
		await transparency.dispatchEvent('change');
		await expect
			.poll(async () =>
				decodeURIComponent((await paint.getAttribute('href'))!.split(',')[1]!).includes('<mask'),
			)
			.toBe(true);
		await host.evaluate((node) => (node as unknown as { undo(): void }).undo());
		await expect(paint).toHaveAttribute('href', original!);
		const type = pane.getByRole('combobox', { name: 'Type', exact: true });
		await expect(type).toHaveValue('rect');
		await stops.getByRole('button', { name: 'Gradient stop 2', exact: true }).click();
		await type.selectOption('linear');
		await expect(host.locator('linearGradient[id$="-s0"]').first()).toBeAttached();
		await expect(pane.getByRole('spinbutton', { name: 'Angle', exact: true })).toBeEnabled();
		await expect(pane.getByRole('spinbutton', { name: 'Position', exact: true })).toHaveValue(
			'100',
		);
		await type.selectOption('rect');
		await expect(paint).toHaveAttribute('href', original!);
		for (const [index, label] of [
			[0, 'From Center'],
			[4, 'From Top Left Corner'],
			[5, 'From Top Right Corner'],
			[6, 'From Bottom Left Corner'],
			[7, 'From Bottom Right Corner'],
		] as const) {
			await pane.getByRole('button', { name: 'Direction', exact: true }).click();
			const popup = host.getByRole('dialog', { name: 'Direction', exact: true });
			await expect(popup.getByRole('button')).toHaveCount(5);
			await popup.getByRole('button', { name: label, exact: true }).click();
			await expect(popup).toBeHidden();
			await expect(pane.getByRole('spinbutton', { name: 'Position', exact: true })).toHaveValue(
				'100',
			);
			const bytes = await host.evaluate(async (node) =>
				Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
			);
			const saved = (await loadXlsx(new Uint8Array(bytes))).sheets[0]!.drawings[0]!;
			const expected = parseDrawingFill(
				parseXml(`<a:spPr xmlns:a="${NS.a}">${native.cases[index]!.fillXml}</a:spPr>`)
					.documentElement,
			)!;
			if (
				saved.kind !== 'chart' ||
				saved.series[0]!.fill?.kind !== 'gradient' ||
				expected.kind !== 'gradient'
			)
				throw new Error('Expected native geometry');
			expect(saved.series[0]!.fill.fillToRect).toEqual(expected.fillToRect);
			expect(saved.series[0]!.fill.tileRect).toEqual(expected.tileRect);
			expect(saved.series[0]!.fill.stops).toEqual(preview.series[0]!.fill.stops);
		}
		await host.evaluate((node) => ((node as unknown as { readOnly: boolean }).readOnly = true));
		await expect(type).toBeDisabled();
		await expect(pane.getByRole('button', { name: 'Direction', exact: true })).toBeDisabled();
		expect(errors).toEqual([]);
	});
