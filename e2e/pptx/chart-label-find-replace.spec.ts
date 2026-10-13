/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright owns the test globals */
import { readFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { createChartLabelDeck } from '../../src/core/pptx/__tests__/helpers/chart-label-deck';
import { savePptxViaBackstage } from './save-pptx';
import { slideStage } from './support/deck';
import { armKeyboard, pressShortcut } from './support/keyboard';

test('explicit chart data scope replaces series and categories with undo and coherent workbook export', async ({
	page,
}) => {
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'chart-labels.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: Buffer.from(await createChartLabelDeck()),
	});
	await expect(slideStage(page)).toBeVisible();
	await armKeyboard(page);
	await page.keyboard.press('Control+h');
	const query = page.locator('input[placeholder*="Find" i]').first();
	const replacement = page.locator('input[placeholder*="Replace with" i]').first();
	const replaceAll = page.getByRole('button', { name: /^Replace (?:All|all matches)$/u });
	await query.fill('Dataset');
	await expect(replaceAll).toBeDisabled();
	await page
		.getByRole('checkbox', { name: 'Include chart series and category labels', exact: true })
		.check();
	await expect(replaceAll).toBeEnabled();
	await replacement.fill('Renamed dataset');
	await replaceAll.click();
	await query.fill('Q1');
	await replacement.fill('First quarter');
	await replaceAll.click();
	await expect(page.locator('[data-pptx-viewport]')).toContainText('First quarter');
	await page.getByRole('button', { name: /^Close find (?:and replace|bar)$/u }).click();
	await pressShortcut(page, 'Control+z');
	await expect(page.locator('[data-pptx-viewport]')).toContainText('Q1');
	await pressShortcut(page, 'Control+y');
	await expect(page.locator('[data-pptx-viewport]')).toContainText('First quarter');
	const download = await savePptxViaBackstage(page);
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	const chart = await zip.file('ppt/charts/chart1.xml')!.async('string');
	const workbook = await JSZip.loadAsync(
		await zip.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array'),
	);
	const sheet = await workbook.file('xl/worksheets/sheet1.xml')!.async('string');
	for (const xml of [chart, sheet]) {
		expect(xml).toContain('Renamed dataset');
		expect(xml).toContain('First quarter');
	}
	expect(chart).toContain('Sheet1!$B$1');
	expect(chart).toContain('Sheet1!$A$2:$A$3');
	expect(sheet).toContain('001 untouched ');
	expect(sheet).toContain('s="7"');
});

test('formula-dependent chart sources remain searchable with explicit replacement rejection', async ({
	page,
}) => {
	const source = await createChartLabelDeck({ formula: true });
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'formula-chart-labels.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: Buffer.from(source),
	});
	await expect(slideStage(page)).toBeVisible();
	await armKeyboard(page);
	await page.keyboard.press('Control+h');
	await page
		.getByRole('checkbox', { name: 'Include chart series and category labels', exact: true })
		.check();
	await page.locator('input[placeholder*="Find" i]').first().fill('Dataset');
	await expect(
		page.getByText(
			'Linked or unsupported chart text can be found, but must be edited in PowerPoint or its source workbook.',
			{ exact: true },
		),
	).toBeVisible();
	await expect(page.getByRole('button', { name: /^Replace(?: current match)?$/u })).toBeDisabled();
	await expect(page.getByRole('button', { name: /^Replace (?:All|all matches)$/u })).toBeDisabled();
	await page.getByRole('button', { name: /^Close find (?:and replace|bar)$/u }).click();
	const download = await savePptxViaBackstage(page),
		saved = await JSZip.loadAsync(await readFile((await download.path())!)),
		original = await JSZip.loadAsync(source);
	expect(await saved.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array')).toEqual(
		await original.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array'),
	);
});
