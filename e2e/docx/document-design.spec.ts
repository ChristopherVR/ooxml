import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { demoReady, newDocument, reopenDocument, saveButton } from './helpers';

test('Word Design previews, applies, undoes and reopens design choices', async ({ page }) => {
	await page.goto('/?framework=vanilla');
	await demoReady(page);
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const surface = editor.locator('.ProseMirror').first();
	await surface.fill('Theme sample');
	const requestsDesign = () =>
		page.evaluate(() =>
			performance.getEntriesByType('resource').some((entry) => /design-gallery-/.test(entry.name)),
		);
	expect(await requestsDesign()).toBe(false);
	await editor.getByRole('tab', { name: 'Design', exact: true }).click();
	expect(await requestsDesign()).toBe(false);
	await expect(editor.getByRole('button', { name: 'Set as Default', exact: true })).toBeDisabled();

	const choose = async (name: string, choice: string) => {
		await editor.getByRole('button', { name, exact: true }).click();
		const menu = editor.getByRole('menu', { name, exact: true });
		await menu.getByRole('menuitemradio', { name: choice, exact: true }).click();
	};
	await choose('Themes', 'Facet');
	await expect.poll(requestsDesign).toBe(true);
	await choose('Fonts', 'Georgia');
	await expect(surface.locator('p')).toHaveCSS('font-family', /Georgia/);
	await choose('Colors', 'Grayscale');
	await choose('Effects', 'Top Shadow');

	const spacing = editor.getByRole('button', { name: 'Paragraph spacing', exact: true });
	await spacing.focus();
	await spacing.press('Enter');
	const double = editor.getByRole('menuitemradio', { name: 'Double', exact: true });
	await double.press('End');
	await page.keyboard.press('ArrowUp');
	await expect(double).toBeFocused();
	await double.press('Enter');
	await expect(surface.locator('p')).toHaveCSS('line-height', '35.2px');

	const undo = editor.locator('.qat').getByRole('button', { name: 'Undo', exact: true });
	const redo = editor.locator('.qat').getByRole('button', { name: 'Redo', exact: true });
	await undo.click();
	await expect(surface.locator('p')).not.toHaveCSS('line-height', '35.2px');
	await redo.click();
	await expect(surface.locator('p')).toHaveCSS('line-height', '35.2px');

	const downloaded = page.waitForEvent('download');
	await saveButton(page).click();
	const download = await downloaded;
	const path = await download.path();
	if (!path) throw new Error('Saved document has no download path');
	const bytes = await readFile(path);
	const zip = await JSZip.loadAsync(bytes);
	const theme = await zip.file('word/theme/theme1.xml')?.async('string');
	expect(theme).toContain('name="Grayscale"');
	expect(theme).toContain('typeface="Georgia"');
	expect(theme).toContain('name="Top Shadow"');
	await reopenDocument(page, 'Design test.docx', bytes);
	await expect(surface).toContainText('Theme sample');
	await expect(surface.locator('p')).toHaveCSS('font-family', /Georgia/);
	await expect(surface.locator('p')).toHaveCSS('line-height', '35.2px');
});

test('Word exports an edited theme and imports it into a new document', async ({ page }) => {
	await page.goto('/?framework=vanilla');
	await demoReady(page);
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const button = (name: string) => editor.getByRole('button', { name, exact: true });
	await editor.getByRole('tab', { name: 'Design', exact: true }).click();
	await button('Themes').click();
	await editor.getByRole('menuitemradio', { name: 'Wisp', exact: true }).click();
	await button('Effects').click();
	await editor.getByRole('menuitemradio', { name: 'Glow Edge', exact: true }).click();
	await button('Themes').click();
	await editor.getByRole('menuitem', { name: 'Save Current Theme', exact: true }).click();
	const dialog = editor.locator('office-ui-dialog.dve-design-dialog');
	await dialog.getByRole('textbox', { name: 'Name', exact: true }).fill('Research');
	const download = page.waitForEvent('download');
	await dialog.getByRole('button', { name: 'Save', exact: true }).click();
	const saved = await download;
	expect(saved.suggestedFilename()).toBe('Research.thmx');
	const path = await saved.path();
	if (!path) throw new Error('Saved theme has no download path');
	const bytes = await readFile(path);
	const zip = await JSZip.loadAsync(bytes);
	const part = await zip.file('theme/theme/theme1.xml')?.async('string');
	expect(part).toContain('name="Glow Edge"');
	expect(part).toContain('<a:glow');
	await newDocument(page);
	await editor
		.locator('office-ui-message-dialog')
		.getByRole('button', { name: 'Discard', exact: true })
		.click();
	await editor.getByRole('tab', { name: 'Design', exact: true }).click();
	await button('Themes').click();
	const chooser = page.waitForEvent('filechooser');
	await editor.getByRole('menuitem', { name: 'Browse for Themes', exact: true }).click();
	await (
		await chooser
	).setFiles({ name: 'Research.thmx', mimeType: 'application/vnd.ms-officetheme', buffer: bytes });
	await expect
		.poll(() =>
			editor.evaluate((element) => {
				if ('documentModel' in element && typeof element.documentModel === 'object')
					return JSON.stringify(element.documentModel);
			}),
		)
		.toContain('Research');
	await button('Effects').click();
	await expect(
		editor.getByRole('menuitemradio', { name: 'Glow Edge', exact: true }),
	).toHaveAttribute('aria-checked', 'true');
});
