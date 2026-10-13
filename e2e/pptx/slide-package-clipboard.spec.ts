/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright owns the test globals */
import { readFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { createSlideClipboardDeck } from '../../src/core/pptx/__tests__/helpers/slide-clipboard-deck';
import { savePptxViaBackstage } from './save-pptx';
import { resetTabSession, slideStage } from './support/deck';
import { pressShortcut } from './support/keyboard';

const mimeType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
test('system clipboard pastes selected slides into another deck with assets, independent themes and one undo step', async ({
	page,
	context,
}) => {
	await context.grantPermissions(['clipboard-read', 'clipboard-write']);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'source-clipboard.pptx',
		mimeType,
		buffer: Buffer.from(await createSlideClipboardDeck(true)),
	});
	await expect(slideStage(page)).toBeVisible();
	await page
		.getByRole('button', { name: /slide sorter/iu })
		.last()
		.click();
	const tiles = page.locator('[data-pptx-chrome="sorter-tile"]');
	await tiles.nth(0).click();
	await page.getByRole('button', { name: 'Copy slides to system clipboard', exact: true }).click();
	await expect(page.getByText('Slides copied', { exact: true })).toBeVisible();
	const preview = await page.evaluate(() => navigator.clipboard.readText());
	expect(preview).toContain('Source copied');
	expect(preview).not.toContain('Private unselected');
	await page.keyboard.press('Escape');
	await expect(tiles).toHaveCount(0);
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'destination-clipboard.pptx',
		mimeType,
		buffer: Buffer.from(await createSlideClipboardDeck(false)),
	});
	await expect(slideStage(page)).toContainText('Destination first');
	await page
		.getByRole('button', { name: /slide sorter/iu })
		.last()
		.click();
	await expect(tiles).toHaveCount(2);
	await tiles.nth(0).click();
	await page
		.getByRole('button', { name: 'Paste slides from system clipboard', exact: true })
		.click();
	await expect(tiles).toHaveCount(3);
	await expect(tiles.nth(1)).toContainText('Source copied');
	await expect(tiles.nth(2)).toContainText('Destination last');
	await expect(tiles.nth(1).locator('img[src]').first()).toBeVisible();
	await expect
		.poll(() =>
			tiles
				.nth(1)
				.locator('img[src]')
				.first()
				.evaluate((image) => (image as HTMLImageElement).naturalWidth),
		)
		.toBe(1);
	await expect
		.poll(() =>
			tiles
				.nth(1)
				.getByText('Source copied', { exact: true })
				.first()
				.evaluate((node) => getComputedStyle(node).color),
		)
		.toBe('rgb(171, 205, 239)');
	await pressShortcut(page, 'Control+z');
	await expect(tiles).toHaveCount(2);
	await pressShortcut(page, 'Control+y');
	await expect(tiles).toHaveCount(3);
	await page.keyboard.press('Escape');
	await expect(tiles).toHaveCount(0);
	const download = await savePptxViaBackstage(page);
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	expect(await zip.file('ppt/theme/theme1.xml')!.async('string')).toContain('112233');
	const themeFiles = Object.values(zip.files).filter((file) =>
		/^ppt\/theme\/theme\d+\.xml$/u.test(file.name),
	);
	expect(
		(await Promise.all(themeFiles.map((file) => file.async('string')))).some((xml) =>
			xml.includes('ABCDEF'),
		),
	).toBe(true);
	const noteFiles = Object.values(zip.files).filter((file) =>
		/^ppt\/notesSlides\/notesSlide\d+\.xml$/u.test(file.name),
	);
	expect(
		(await Promise.all(noteFiles.map((file) => file.async('string')))).some(
			(xml) => xml.includes('Source rich speaker notes') && xml.includes('b="1"'),
		),
	).toBe(true);
	expect(Object.values(zip.files).some((file) => /^ppt\/media\/.*\.png$/u.test(file.name))).toBe(
		true,
	);
});

test('foreign system clipboard content is reported instead of pasting retained local slides', async ({
	page,
	context,
}) => {
	await context.grantPermissions(['clipboard-read', 'clipboard-write']);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'destination-clipboard.pptx',
		mimeType,
		buffer: Buffer.from(await createSlideClipboardDeck(false)),
	});
	await expect(slideStage(page)).toBeVisible();
	await page
		.getByRole('button', { name: /slide sorter/iu })
		.last()
		.click();
	await page.evaluate(() => navigator.clipboard.writeText('foreign clipboard text'));
	await page
		.getByRole('button', { name: 'Paste slides from system clipboard', exact: true })
		.click();
	await expect(
		page.getByText('The clipboard does not contain copied PowerPoint slides', { exact: true }),
	).toBeVisible();
	await expect(page.locator('[data-pptx-chrome="sorter-tile"]')).toHaveCount(2);
});
