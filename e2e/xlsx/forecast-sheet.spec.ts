import { expect, test } from '@playwright/test';
import {
	clickRibbonControl,
	editor,
	FRAMEWORKS,
	formulaBarValue,
	goToCell,
	newWorkbook,
	pageErrors,
	ribbon,
	sheetTabs,
	typeInActiveCell,
} from './helpers';

for (const framework of FRAMEWORKS)
	test(`${framework}: Forecast Sheet creates a sheet with formulas and a chart`, async ({
		page,
	}) => {
		const errors = pageErrors(page);
		await newWorkbook(page, framework);
		await goToCell(page, 'A1');
		for (const text of ['Month', '1', '2', '3', '4', '5', '6', '7', '8'])
			await typeInActiveCell(page, text);
		await goToCell(page, 'B1');
		for (const text of ['Sales', '10', '12', '13', '15', '16', '18', '19', '21'])
			await typeInActiveCell(page, text);
		await goToCell(page, 'A1');
		await page.keyboard.press('Control+Shift+ArrowDown');
		await page.keyboard.press('Shift+ArrowRight');

		await ribbon(page).getByRole('tab', { name: 'Data', exact: true }).click();
		await clickRibbonControl(page, ribbon(page).locator('[data-command="data.forecast-sheet"]'));
		const dialog = editor(page).locator('[data-dialog="forecast-sheet"]');
		await expect(dialog).toBeVisible();
		await expect(dialog.getByLabel('Timeline range:')).toHaveValue('$A$1:$A$9');
		await expect(dialog.getByLabel('Values range:')).toHaveValue('$B$1:$B$9');
		await dialog.getByLabel('Forecast end:').fill('12');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(dialog).toBeHidden();

		// The forecast lands on a new, active sheet: history, then FORECAST.ETS rows with bounds.
		await expect(sheetTabs(page).getByRole('tab')).toHaveCount(2);
		await goToCell(page, 'C10');
		await expect.poll(() => formulaBarValue(page)).toMatch(/^=FORECAST\.ETS\(A10,/);
		await goToCell(page, 'D10');
		await expect.poll(() => formulaBarValue(page)).toContain('FORECAST.ETS.CONFINT');
		await expect(editor(page).locator('[part="grid"] .xg-obj svg').first()).toBeVisible();
		expect(errors).toEqual([]);
	});
