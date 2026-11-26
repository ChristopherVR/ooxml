import { expect, test } from '@playwright/test';
import { dialogByHeading } from './helpers';
import { openShapes, unfoldArrange } from './word-arrange-support';

test('Layout KeyTips open the Position gallery and Alt Text restores keyboard focus', async ({
	page,
}) => {
	const editor = await openShapes(page);
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await editor.locator('.dve-print-float').nth(2).click({ force: true });
	await editor.evaluate((element) => {
		const story = element.shadowRoot!.querySelector('.ProseMirror')!;
		for (const type of ['keydown', 'keyup'])
			story.dispatchEvent(new KeyboardEvent(type, { key: 'Alt', bubbles: true, composed: true }));
	});
	await page.keyboard.press('p');
	const position = editor.getByRole('combobox', { name: 'Position', exact: true });
	await expect(editor.locator('.dve-keytip-command').filter({ hasText: /^PO$/ })).toBeVisible();
	await page.keyboard.press('p');
	await page.keyboard.press('o');
	const gallery = editor.getByRole('menu', { name: 'Position' });
	const first = gallery.getByRole('menuitemradio').first();
	await expect(first).toBeFocused();
	await page.keyboard.press('ArrowDown');
	await expect(gallery.getByRole('menuitemradio', { name: /Top Left/ })).toBeFocused();
	await page.keyboard.press('Escape');
	await expect(position).toBeFocused();
	await page.keyboard.press('Tab');
	await expect(editor.getByRole('combobox', { name: 'Wrap text', exact: true })).toBeFocused();

	const altText = editor.getByRole('button', { name: 'Alt Text', exact: true });
	await altText.press('Enter');
	const dialog = dialogByHeading(editor, 'Alt Text');
	await expect(dialog.getByRole('textbox', { name: 'Description' })).toBeFocused();
	await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
	await expect(dialog).toBeHidden();
	await expect(altText).toBeFocused();
});

test('French labels reach Arrange menus, disabled reasons and the Selection Pane', async ({
	page,
}) => {
	const editor = await openShapes(page);
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await editor.locator('.dve-print-float').nth(2).click({ force: true });
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await unfoldArrange(editor, 'Layout');
	await editor.evaluate((element) => element.setAttribute('locale', 'fr'));
	const group = editor.locator('select[data-arrange="group"]:visible');
	await expect(group).toHaveAccessibleDescription(
		'Sélectionnez au moins deux formes flottantes dans un paragraphe.',
	);
	await editor.locator('select[data-arrange="position"]:visible').press('Enter');
	await expect(editor.getByRole('menuitemradio', { name: /En bas à droite/ })).toBeVisible();
	await page.keyboard.press('Escape');
	await editor.getByRole('button', { name: 'Volet Sélection', exact: true }).click();
	const pane = editor.getByRole('complementary', { name: 'Volet Sélection' });
	await pane.getByRole('button', { name: 'Oval 3' }).press('F2');
	await expect(pane.getByRole('textbox', { name: 'Nom du dessin' })).toBeFocused();
});
