import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { createWorkbook, createEditSession, saveXlsx } from 'ooxml-core/xlsx';
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
