import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { newDocument, saveButton, dialogByHeading } from './helpers';

test('Insert > Text Box inserts, edits and exports an inline text box', async ({ page }) => {
	await page.setViewportSize({ width: 1700, height: 720 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('Before ');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await editor.getByRole('button', { name: 'Text box', exact: true }).click();
	const dialog = dialogByHeading(editor, 'Text box');
	await expect(dialog).toBeVisible();
	await dialog.getByRole('textbox', { name: 'Text', exact: true }).fill('Hello box\nSecond line');
	await dialog.getByRole('spinbutton', { name: 'Width (inches)', exact: true }).fill('0');
	await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeDisabled();
	await dialog.getByRole('spinbutton', { name: 'Width (inches)', exact: true }).fill('2');
	expect(
		await dialog
			.getByRole('button', { name: 'OK', exact: true })
			.evaluate((button) => button.getBoundingClientRect().bottom <= window.innerHeight),
	).toBe(true);
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	// Drawn like any shape: its outline, and its lines as paragraphs over it.
	const box = editor.locator('.ProseMirror .dve-shape');
	await expect(box).toHaveCount(1);
	await expect(box.locator('.dve-text-box-content p')).toHaveText(['Hello box', 'Second line']);
	await expect(box.locator('svg path')).toHaveAttribute('stroke', /^#000000$/i);
	await expect(editor.locator('.ProseMirror .dve-image-placeholder')).toHaveCount(0);
	// Selecting the box shows Shape Format, as in Word; Insert > Text Box then edits the box
	// instead of inserting another.
	await box.click();
	await expect(editor.getByRole('tab', { name: 'Shape Format', exact: true })).toBeVisible();
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await editor.getByRole('button', { name: 'Text box', exact: true }).click();
	await expect(dialog.getByRole('textbox', { name: 'Text', exact: true })).toHaveValue(
		'Hello box\nSecond line',
	);
	await dialog.getByRole('textbox', { name: 'Text', exact: true }).fill('Edited');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(box).toHaveCount(1);
	await expect(box.locator('.dve-text-box-content p')).toHaveText(['Edited']);
	// Shape Format > Add Text edits the same box in the Text Box dialog.
	await editor.getByRole('tab', { name: 'Shape Format', exact: true }).click();
	await editor.getByRole('button', { name: 'Add text', exact: true }).click();
	await expect(dialog.getByRole('textbox', { name: 'Text', exact: true })).toHaveValue('Edited');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	const downloadPromise = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await downloadPromise).path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toContain('txBox="1"');
	expect(xml).toContain('<w:t>Edited</w:t>');
	expect(xml.match(/<wp:inline/g)).toHaveLength(1);
});
