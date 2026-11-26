import { expect, test } from '@playwright/test';
import { openSample } from './helpers';

test('File > Customize Ribbon hides and restores commands', async ({ page }) => {
	await page.setViewportSize({ width: 1700, height: 800 });
	await openSample(page);
	const editor = page.locator('docx-editor');
	const events = await editor.evaluate((element) => {
		(window as unknown as { seen: string[][] }).seen = [];
		element.addEventListener('ribbon-customize', (event) =>
			(window as unknown as { seen: string[][] }).seen.push([
				...(event as CustomEvent<string[]>).detail,
			]),
		);
		return true;
	});
	expect(events).toBe(true);
	const dialog = editor.locator('office-ui-options-dialog');
	const openCustomize = async () => {
		await editor.locator('.dve-ribbon .file').click();
		await editor.locator('.dve-backstage .item', { hasText: 'Options' }).click();
		await dialog.getByRole('tab', { name: 'Customize Ribbon' }).click();
	};
	await openCustomize();
	await editor
		.locator('.dve-customize-row', { hasText: /^Bold$/ })
		.locator('input')
		.uncheck();
	await editor
		.locator('.dve-customize-row', { hasText: /^Select Recipients$/ })
		.locator('input')
		.uncheck();
	await dialog.locator('[data-action="ok"]').click();
	await expect(editor.getByRole('button', { name: 'Bold', exact: true })).toBeHidden();
	expect(
		await page.evaluate(() => (window as unknown as { seen: string[][] }).seen.at(-1)),
	).toEqual(['bold', 'select-recipients']);
	await openCustomize();
	await editor.getByRole('button', { name: 'Reset all customizations' }).click();
	await dialog.locator('[data-action="ok"]').click();
	await expect(editor.getByRole('button', { name: 'Bold', exact: true })).toBeVisible();
	await editor.getByRole('tab', { name: 'Mailings' }).click();
	await expect(editor.getByRole('button', { name: 'Select Recipients' })).toBeVisible();
});

test('Tell me opens a Mailings command directly from Home', async ({ page }) => {
	await page.setViewportSize({ width: 1700, height: 800 });
	await openSample(page);
	const editor = page.locator('docx-editor');
	await expect(editor.getByRole('tab', { name: 'Home' })).toHaveAttribute('aria-selected', 'true');
	await editor.getByRole('searchbox', { name: 'Tell me what you want to do' }).fill('envelopes');
	await editor.getByRole('option', { name: 'Envelopes' }).click();
	await expect(editor.getByRole('tab', { name: 'Mailings' })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	await expect(editor.getByRole('textbox', { name: 'Delivery address' })).toBeFocused();
});
