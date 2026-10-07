import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { createWorkbook, createEditSession, loadXlsx, saveXlsx } from 'ooxml-core/xlsx';
import native from '../../src/core/chart/excel-chart-styles.json' with { type: 'json' };
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';

async function openNativeStyle(page: Page, style: number, framework = 'vanilla') {
	const sample = native.cases.find((sample) => sample.style === style)!;
	const book = createWorkbook();
	book.theme.majorFont = 'Aptos Display';
	book.theme.minorFont = 'Aptos Narrow';
	book.theme.colors = [
		'lt1',
		'dk1',
		'lt2',
		'dk2',
		'accent1',
		'accent2',
		'accent3',
		'accent4',
		'accent5',
		'accent6',
		'hlink',
		'folHlink',
	].map((slot) => (native.scheme as Record<string, string>)[slot]!.slice(1));
	createEditSession(book).addChart(0, {
		chartType: 'column',
		series: [],
		showLegend: true,
		anchor: {
			from: { row: 1, col: 4, rowOffset: 0, colOffset: 0 },
			ext: { cx: 6096000, cy: 3810000 },
		},
	});
	const zip = await JSZip.loadAsync(await saveXlsx(book));
	for (const [name, xml] of Object.entries(sample.parts)) zip.file(name, xml);
	await openLanding(page, framework);
	await page.locator('#landing-file').setInputFiles({
		name: `native-${style}.xlsx`,
		mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		buffer: await zip.generateAsync({ type: 'nodebuffer' }),
	});
	const chart = editor(page).locator('.xg-chart svg').first();
	await expect(chart).toBeVisible();
	return chart;
}

for (const framework of FRAMEWORKS)
	test(`native series spacing edits in ${framework}`, async ({ page }) => {
		const errors = pageErrors(page);
		const chart = await openNativeStyle(page, 209, framework);
		await page.getByRole('img', { name: 'Native style', exact: true }).first().click();
		await editor(page).getByRole('tab', { name: 'Chart Design', exact: true }).click();
		await editor(page).getByRole('button', { name: 'Format Data Series', exact: true }).click();
		const pane = editor(page).getByRole('complementary', { name: 'Format Data Series' });
		const gap = pane.getByRole('spinbutton', { name: 'Gap Width', exact: true });
		const overlap = pane.getByRole('spinbutton', { name: 'Series Overlap', exact: true });
		await expect(gap).toHaveValue('100');
		await expect(overlap).toHaveValue('-24');
		await gap.fill('5');
		await gap.press('Tab');
		await overlap.fill('23');
		await overlap.press('Tab');
		const first = chart.locator('g[data-chart-series="0"][data-chart-point="0"] rect');
		await expect
			.poll(() =>
				chart.evaluate((svg) => {
					const a = svg.querySelector('g[data-chart-series="0"][data-chart-point="0"] rect')!;
					const b = svg.querySelector('g[data-chart-series="0"][data-chart-point="1"] rect')!;
					return (
						Number(a.getAttribute('width')) /
						(Number(b.getAttribute('x')) - Number(a.getAttribute('x')))
					);
				}),
			)
			.toBeCloseTo(1 / 1.82, 3);
		const bytes = await editor(page).evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		const saved = await loadXlsx(new Uint8Array(bytes));
		expect(saved.sheets[0]!.drawings[0]).toMatchObject({ barGapWidth: 5, barOverlap: 23 });
		await editor(page).evaluate((node) => (node as unknown as { undo(): void }).undo());
		await expect(overlap).toHaveValue('-24');
		await editor(page).evaluate((node) => (node as unknown as { redo(): void }).redo());
		await expect(overlap).toHaveValue('23');
		await editor(page).evaluate((node) => {
			(node as unknown as { readOnly: boolean }).readOnly = true;
		});
		await expect(gap).toBeDisabled();
		await editor(page).evaluate((node) => {
			(node as unknown as { readOnly: boolean }).readOnly = false;
		});
		await expect(gap).toBeEnabled();
		await pane.getByRole('button', { name: 'Close', exact: true }).click();
		await expect(pane).toBeHidden();
		await first.dblclick();
		await expect(pane).toBeVisible();
		await overlap.press('Escape');
		await expect(pane).toBeHidden();
		expect(errors).toEqual([]);
	});

for (const framework of FRAMEWORKS)
	test(`native title and legend typography render in ${framework}`, async ({ page }) => {
		const errors = pageErrors(page);
		const chart = await openNativeStyle(page, 212, framework);
		const title = chart.locator('text').filter({ hasText: 'Native style' });
		await expect(title).toHaveAttribute('font-family', /Aptos Display.*sans-serif/);
		expect(Number(await title.getAttribute('font-size'))).toBeCloseTo((20 * 4) / 3);
		await expect(title).toHaveAttribute('fill', '#595959');
		await expect(chart.locator('text').filter({ hasText: /^Sales$/ })).toHaveAttribute(
			'font-size',
			'12',
		);
		await page.getByRole('img', { name: 'Native style', exact: true }).first().click();
		await editor(page).getByRole('tab', { name: 'Chart Design', exact: true }).click();
		await expect(
			editor(page).getByRole('tab', { name: 'Chart Design', exact: true }),
		).toHaveAttribute('aria-selected', 'true');
		expect(errors).toEqual([]);
	});

test('native hidden value axes suppress labels while retaining categories', async ({ page }) => {
	const chart = await openNativeStyle(page, 202);
	await expect(chart.locator('text').filter({ hasText: /^M2$/ })).toBeVisible();
	await expect(chart.locator('text').filter({ hasText: /^0$/ })).toHaveCount(0);
});

test('native dark chart styles retain gradient backgrounds and contrasting text', async ({
	page,
}) => {
	const chart = await openNativeStyle(page, 209);
	await expect(chart.locator('radialGradient')).toHaveCount(1);
	await expect(chart.locator('radialGradient stop')).toHaveCount(2);
	await expect(chart.locator('text').filter({ hasText: 'Native style' })).toHaveAttribute(
		'fill',
		'#F2F2F2',
	);
	await expect(chart.locator('rect').first()).toHaveAttribute('fill', /^url\(#xlsx-chart-.*\)$/);
	await expect(chart.locator('line[stroke="rgba(242,242,242,0.1)"]')).toHaveCount(7);
	await expect(chart.locator('linearGradient')).toHaveCount(2);
	await expect(chart.locator('linearGradient stop')).toHaveCount(6);
	await expect(chart.locator('feDropShadow')).toHaveCount(3);
	await expect(chart.locator('feDropShadow[stdDeviation="3"][dy="2"]')).toHaveCount(2);
	const firstBar = chart.locator('g[data-chart-series="0"][data-chart-point="0"] rect');
	const nextBar = chart.locator('g[data-chart-series="0"][data-chart-point="1"] rect');
	const pitch = Number(await nextBar.getAttribute('x')) - Number(await firstBar.getAttribute('x'));
	expect(Number(await firstBar.getAttribute('width')) / pitch).toBeCloseTo(1 / 3.24, 3);
	await expect(chart.locator('text').filter({ hasText: 'Native style' })).toHaveAttribute(
		'filter',
		/title-text-shadow/,
	);
	await expect(chart.locator('linearGradient stop').first()).toHaveAttribute(
		'stop-color',
		'#497491',
	);
	await expect(chart.locator('rect[fill*="-s0)"]')).toHaveCount(5);
	await page.getByRole('img', { name: 'Native style', exact: true }).first().click();
	await editor(page).getByRole('tab', { name: 'Chart Design', exact: true }).click();
	await editor(page).getByRole('button', { name: 'Change Colors', exact: true }).click();
	await page.getByRole('button', { name: 'Colorful Palette 3', exact: true }).click();
	await expect(chart.locator('linearGradient stop').first()).toHaveAttribute(
		'stop-color',
		'#ED8256',
	);
	await expect(chart.locator('radialGradient stop')).toHaveCount(2);
});
