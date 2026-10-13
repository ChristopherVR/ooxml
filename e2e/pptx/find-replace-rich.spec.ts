/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright owns the test globals */
import { readFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { createRichFindDeck } from '../../src/core/pptx/__tests__/helpers/find-replace-deck';
import { savePptxViaBackstage } from './save-pptx';
import { slideStage } from './support/deck';
import { armKeyboard, pressShortcut } from './support/keyboard';

test('find replaces text spanning rich runs, grouped text and table cells with undo and saved formatting', async ({
	page,
}) => {
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'rich-find.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: Buffer.from(await createRichFindDeck()),
	});
	await expect(slideStage(page)).toBeVisible();
	await armKeyboard(page);
	await page.keyboard.press('Control+h');
	const query = page.locator('input[placeholder*="Find" i]').first();
	await query.fill('report');
	await expect(
		page.getByText(/^\s*1\s+(?:of|\/)\s+3(?:\s+result\(s\) across 1 slide\(s\))?\s*$/u),
	).toBeVisible();
	await page.locator('input[placeholder*="Replace with" i]').first().fill('summary');
	await page.getByRole('button', { name: /^Replace (?:All|all matches)$/u }).click();
	await query.fill('summary');
	await expect(
		page.getByText(/^\s*1\s+(?:of|\/)\s+3(?:\s+result\(s\) across 1 slide\(s\))?\s*$/u),
	).toBeVisible();
	await expect(page.locator('[data-pptx-viewport]')).toContainText('Annual summary');
	await page.getByRole('button', { name: /^Close find (?:and replace|bar)$/u }).click();
	await pressShortcut(page, 'Control+z');
	await page.keyboard.press('Control+h');
	await query.fill('report');
	await expect(
		page.getByText(/^\s*1\s+(?:of|\/)\s+3(?:\s+result\(s\) across 1 slide\(s\))?\s*$/u),
	).toBeVisible();
	await page.getByRole('button', { name: /^Close find (?:and replace|bar)$/u }).click();
	await pressShortcut(page, 'Control+y');
	const download = await savePptxViaBackstage(page);
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	expect(xml.match(/summary/g)).toHaveLength(3);
	expect(xml).toContain('opaque-cell');
	expect(xml).toContain('u="sng"');
	expect(xml).toContain('rIdKeep');
});

test('speaker notes are an explicit find scope and retain formatting after replace and undo', async ({
	page,
}) => {
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'notes-find.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: Buffer.from(await createRichFindDeck(true)),
	});
	await expect(slideStage(page)).toBeVisible();
	await armKeyboard(page);
	await page.keyboard.press('Control+h');
	const query = page.locator('input[placeholder*="Find" i]').first();
	await query.fill('report');
	const count = (total: number) =>
		page.getByText(new RegExp(`^\\s*1\\s+(?:of|/)\\s+${total}\\s*$`, 'u'));
	await expect(count(3)).toBeVisible();
	await page.getByRole('checkbox', { name: 'Include speaker notes', exact: true }).check();
	await expect(count(4)).toBeVisible();
	await page.locator('input[placeholder*="Replace with" i]').first().fill('summary');
	await page.getByRole('button', { name: /^Replace (?:All|all matches)$/u }).click();
	await query.fill('summary');
	await expect(count(4)).toBeVisible();
	await page.getByRole('button', { name: /^Close find (?:and replace|bar)$/u }).click();
	await pressShortcut(page, 'Control+z');
	await page.keyboard.press('Control+h');
	await query.fill('report');
	await page.getByRole('checkbox', { name: 'Include speaker notes', exact: true }).check();
	await expect(count(4)).toBeVisible();
	await page.getByRole('button', { name: /^Close find (?:and replace|bar)$/u }).click();
	await pressShortcut(page, 'Control+y');
	const download = await savePptxViaBackstage(page);
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	const notesXml = await zip.file('ppt/notesSlides/notesSlide1.xml')!.async('string');
	expect(notesXml).toContain('Speaker summary');
	expect(notesXml).toContain('b="1"');
});
