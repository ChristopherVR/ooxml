import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { newDocument, saveButton, dialogByHeading } from './helpers';

test('Design > Watermark inserts a washed-out picture behind Print Layout pages and saves VML', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1700, height: 720 });
	await page.goto('/?framework=vanilla');
	const PNG = Buffer.from(
		await page.evaluate(() => {
			const canvas = document.createElement('canvas');
			canvas.width = 40;
			canvas.height = 20;
			const context = canvas.getContext('2d')!;
			context.fillStyle = '#2563eb';
			context.fillRect(0, 0, 40, 20);
			return canvas.toDataURL('image/png').split(',')[1]!;
		}),
		'base64',
	);
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('Body text');
	await editor.getByRole('tab', { name: 'Design', exact: true }).click();
	await editor.getByRole('button', { name: 'Watermark', exact: true }).click();
	const dialog = dialogByHeading(editor, 'Printed watermark');
	await expect(dialog).toBeVisible();
	await dialog.getByRole('combobox', { name: 'Setting', exact: true }).selectOption('picture');
	await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeDisabled();
	await dialog
		.locator('input[type="file"]')
		.setInputFiles({ name: 'seal.png', mimeType: 'image/png', buffer: PNG });
	await expect(dialog.getByText('seal.png')).toBeVisible();
	await dialog.getByRole('combobox', { name: 'Scale', exact: true }).selectOption('100');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(dialog).toBeHidden();
	await editor.locator('[role="tab"][data-tab="view"]').click();
	await editor.getByRole('button', { name: 'Print Layout' }).click();
	const mark = editor.locator('.dve-print-page .dve-print-watermark-picture');
	await expect(mark).toHaveCount(1);
	await expect(mark).toHaveAttribute('src', /^blob:|^data:/);
	await expect(mark).toHaveCSS('filter', /contrast/);

	await editor.getByRole('tab', { name: 'Design', exact: true }).click();
	await editor.getByRole('button', { name: 'Watermark', exact: true }).click();
	await expect(dialog.getByRole('combobox', { name: 'Setting', exact: true })).toHaveValue(
		'picture',
	);
	await dialog.getByRole('checkbox', { name: 'Washout', exact: true }).uncheck();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(mark).toHaveCount(1);
	await expect(mark).toHaveCSS('filter', 'none');

	const downloadPromise = page.waitForEvent('download');
	await saveButton(page).click();
	const download = await downloadPromise;
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	const header = Object.keys(zip.files).find((name) => /^word\/header\d*\.xml$/.test(name));
	expect(header).toBeTruthy();
	const xml = await zip.file(header!)!.async('string');
	expect(xml).toMatch(/id="WordPictureWatermark\d+"/);
	expect(xml).toContain('type="#_x0000_t75"');
	expect(xml).not.toContain('gain=');
	const rels = await zip.file(header!.replace('word/', 'word/_rels/') + '.rels')!.async('string');
	const target = /relationships\/image" Target="([^"]+)"/.exec(rels)?.[1];
	expect(target).toBeTruthy();
	expect(await zip.file(`word/${target}`)!.async('nodebuffer')).toEqual(PNG);

	await editor.getByRole('tab', { name: 'Design', exact: true }).click();
	await editor.getByRole('button', { name: 'Watermark', exact: true }).click();
	await dialog.getByRole('combobox', { name: 'Setting', exact: true }).selectOption('none');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(mark).toHaveCount(0);
});
