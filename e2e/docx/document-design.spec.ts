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
	await expect.poll(requestsDesign).toBe(true);
	await expect(editor.getByRole('button', { name: 'Set as Default', exact: true })).toBeDisabled();

	const choose = async (name: string, choice: string) => {
		await editor.getByRole('button', { name, exact: true }).click();
		const menu = editor.getByRole('menu', { name, exact: true });
		await menu.getByRole('menuitemradio', { name: choice, exact: true }).click();
	};
	await choose('Themes', 'Facet');
	await choose('Fonts', 'Georgia');
	await expect(surface.locator('p')).toHaveCSS('font-family', /Georgia/);
	await choose('Colors', 'Grayscale');
	await choose('Effects', 'Strong Shadow');

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
	expect(theme).toContain('name="strong"');
	await reopenDocument(page, 'Design test.docx', bytes);
	await expect(surface).toContainText('Theme sample');
	await expect(surface.locator('p')).toHaveCSS('font-family', /Georgia/);
	await expect(surface.locator('p')).toHaveCSS('line-height', '35.2px');
});
