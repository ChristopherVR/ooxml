import { expect, test } from '@playwright/test';
import { editor, goToCell, grid, newWorkbook, ribbon, typeInActiveCell } from './helpers';

test('create, format, clear and undo sparklines through the ribbon', async ({ page }, info) => {
	await newWorkbook(page);
	await typeInActiveCell(page, '5');
	await typeInActiveCell(page, '-3');
	await typeInActiveCell(page, '8');
	await goToCell(page, 'B1');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await ribbon(page).locator('[data-command="insert.sparkline-line"]').click();
	const create = editor(page).locator('[data-dialog="create-sparklines"]');
	await create.getByLabel('Data Range:', { exact: true }).fill('A1:A3');
	await create.getByRole('button', { name: 'OK', exact: true }).click();
	const spark = grid(page).locator('.xg-c:not([hidden]) .xg-spark');
	await expect(spark).toHaveCount(1);
	await ribbon(page).locator('[data-command="sparkline.settings"]').first().click();
	const settings = editor(page).locator('[data-dialog="sparkline-settings"]');
	await settings.getByLabel('Sparkline Type:', { exact: true }).selectOption('column');
	await settings.getByLabel('High Point', { exact: true }).check();
	await page.screenshot({ path: info.outputPath('sparkline-settings.png') });
	await settings.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(spark.locator('rect')).toHaveCount(3);
	await ribbon(page).getByRole('button', { name: 'Clear Sparklines', exact: true }).click();
	await editor(page)
		.getByRole('menuitem', { name: 'Clear Selected Sparklines', exact: true })
		.click();
	await expect(spark).toHaveCount(0);
	await page.keyboard.press('Control+Z');
	await expect(spark).toHaveCount(1);
	await page.keyboard.press('Control+Y');
	await expect(spark).toHaveCount(0);
	await editor(page).getByRole('button', { name: 'Undo', exact: true }).click();
	await expect(spark).toHaveCount(1);
});
