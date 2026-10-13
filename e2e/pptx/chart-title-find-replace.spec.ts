/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright owns the test globals */
import { readFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { createChartTitleDeck } from '../../src/core/pptx/__tests__/helpers/chart-title-deck';
import { savePptxViaBackstage } from './save-pptx';
import { slideStage } from './support/deck';
import { armKeyboard, pressShortcut } from './support/keyboard';

test('find replaces a rich chart title across formatting runs with undo, redo and preserved export', async ({
	page,
}) => {
	const source = await createChartTitleDeck();
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'chart-title.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: Buffer.from(source),
	});
	await expect(slideStage(page)).toBeVisible();
	await armKeyboard(page);
	await page.keyboard.press('Control+h');
	const query = page.locator('input[placeholder*="Find" i]').first();
	await query.fill('report');
	await expect(
		page
			.getByText(/^\s*1\s+(?:of|\/)\s+1(?:\s+result\(s\) across 1 slide\(s\))?\s*$/u)
			.filter({ visible: true }),
	).toBeVisible();
	await page.locator('input[placeholder*="Replace with" i]').first().fill('summary');
	await page.getByRole('button', { name: /^Replace (?:All|all matches)$/u }).click();
	await query.fill('summary');
	await expect(
		page
			.getByText(/^\s*1\s+(?:of|\/)\s+1(?:\s+result\(s\) across 1 slide\(s\))?\s*$/u)
			.filter({ visible: true }),
	).toBeVisible();
	await expect(page.locator('[data-pptx-viewport]')).toContainText('Annual summary');
	await page.getByRole('button', { name: /^Close find (?:and replace|bar)$/u }).click();
	await pressShortcut(page, 'Control+z');
	await expect(page.locator('[data-pptx-viewport]')).toContainText('Annual report');
	await pressShortcut(page, 'Control+y');
	await expect(page.locator('[data-pptx-viewport]')).toContainText('Annual summary');
	const download = await savePptxViaBackstage(page);
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	const xml = await zip.file('ppt/charts/chart1.xml')!.async('string');
	for (const marker of [
		'Annual summary',
		'lang="fr-FR"',
		'u="sng"',
		'accent2',
		'rIdTitleLink',
		'chart-title-parity',
	])
		expect(xml).toContain(marker);
	const original = await JSZip.loadAsync(source);
	expect(await zip.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array')).toEqual(
		await original.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array'),
	);
});

test('linked chart title matches explain source editing and disable replacement without changing the workbook', async ({
	page,
}) => {
	const source = await createChartTitleDeck(true);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'linked-chart-title.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: Buffer.from(source),
	});
	await expect(slideStage(page)).toBeVisible();
	await armKeyboard(page);
	await page.keyboard.press('Control+h');
	await page.locator('input[placeholder*="Find" i]').first().fill('report');
	await expect(
		page.getByText(
			'Linked or complex chart titles can be found, but must be edited in PowerPoint or their source workbook.',
			{ exact: true },
		),
	).toBeVisible();
	await expect(page.getByRole('button', { name: /^Replace (?:All|all matches)$/u })).toBeDisabled();
	await expect(page.getByRole('button', { name: /^Replace(?: current match)?$/u })).toBeDisabled();
	await page.getByLabel('Include chart titles', { exact: true }).uncheck();
	await expect(page.getByRole('button', { name: /^Replace (?:All|all matches)$/u })).toBeEnabled();
	await page.locator('input[placeholder*="Replace with" i]').first().fill('summary');
	await page.getByRole('button', { name: /^Replace (?:All|all matches)$/u }).click();
	await expect(page.locator('[data-pptx-viewport]')).toContainText('Slide-only summary');
	await expect(page.locator('[data-pptx-viewport]')).toContainText('Annual report');
	await page.getByRole('button', { name: /^Close find (?:and replace|bar)$/u }).click();
	const download = await savePptxViaBackstage(page),
		zip = await JSZip.loadAsync(await readFile((await download.path())!)),
		original = await JSZip.loadAsync(source);
	expect(await zip.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array')).toEqual(
		await original.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array'),
	);
	expect(await zip.file('ppt/charts/chart1.xml')!.async('string')).toContain('Sheet1!$A$1');
});
