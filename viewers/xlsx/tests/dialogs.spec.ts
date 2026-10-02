import { expect, test } from '@playwright/test';
import { editor, goToCell, grid, newWorkbook, typeInActiveCell } from './helpers';

const dialog = (page: import('@playwright/test').Page, name: string) =>
	editor(page).locator(`[data-dialog="${name}"]`);

test('Find and Replace replaces every match', async ({ page }) => {
	await newWorkbook(page);
	await typeInActiveCell(page, 'apple pie');
	await typeInActiveCell(page, 'apple tart');
	await goToCell(page, 'A1');
	await page.keyboard.press('Control+H');
	const box = dialog(page, 'find-replace');
	await expect(box).toBeVisible();
	await box.getByLabel('Find what:').fill('apple');
	await box.getByLabel('Replace with:').fill('pear');
	await box.getByRole('button', { name: 'Replace All' }).click();
	await expect(grid(page)).toContainText('pear pie');
	await expect(grid(page)).toContainText('pear tart');
	await box.getByRole('button', { name: 'Close' }).last().click();
	await expect(box).toBeHidden();
});

test('Format Cells (Ctrl+1) applies a number format from its Number tab', async ({ page }) => {
	await newWorkbook(page);
	await typeInActiveCell(page, '0.25');
	await goToCell(page, 'A1');
	await page.keyboard.press('Control+1');
	const box = dialog(page, 'format-cells');
	await expect(box).toBeVisible();
	await box.getByRole('option', { name: 'Percentage' }).click();
	await box.getByRole('button', { name: 'OK' }).click();
	await expect(box).toBeHidden();
	await expect(grid(page)).toContainText('25.00%');
});
