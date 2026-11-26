import { expect, test } from '@playwright/test';
import {
	editor,
	goToCell,
	grid,
	newWorkbook,
	openLanding,
	ribbon,
	typeInActiveCell,
} from './helpers';

const PNG = Buffer.from(
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGNoTgkFAAKqAT2gS4eLAAAAAElFTkSuQmCC',
	'base64',
);

test('inserts a cell picture and blocks IMAGE without a host fetcher', async ({ page }) => {
	await openLanding(page);
	await newWorkbook(page);
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	const pictures = editor(page).getByRole('button', { name: 'Pictures', exact: true });
	if (!(await pictures.isVisible()))
		await ribbon(page).getByRole('button', { name: 'Illustrations', exact: true }).click();
	await pictures.click();
	const chooser = page.waitForEvent('filechooser');
	await editor(page).getByText('Place in Cell', { exact: true }).click();
	await (await chooser).setFiles({ name: 'pixel.png', mimeType: 'image/png', buffer: PNG });
	const picture = grid(page).locator('.xg-c > img');
	await expect(picture).toBeVisible();
	await expect(picture).toHaveCSS('object-fit', 'contain');
	await expect
		.poll(() =>
			picture.evaluate((image) => image instanceof HTMLImageElement && image.naturalWidth),
		)
		.toBe(1);
	await grid(page).focus();
	await grid(page).press('Shift+F10');
	await editor(page).getByText('Alt Text', { exact: true }).last().click();
	await editor(page).getByRole('textbox', { name: 'Alt Text' }).fill('Small square');
	await editor(page).getByRole('button', { name: 'OK', exact: true }).click();
	await expect(picture).toHaveAttribute('alt', 'Small square');
	await goToCell(page, 'B1');
	await typeInActiveCell(page, '=IMAGE("https://example.test/picture.png")');
	await expect(
		grid(page)
			.locator('.xg-tx')
			.filter({ hasText: /^#BLOCKED!$/ }),
	).toBeVisible();
});
