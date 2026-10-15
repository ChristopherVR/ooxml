import { expect, test, type Page } from '@playwright/test';
import {
	editor,
	formulaBarValue,
	goToCell,
	grid,
	newWorkbook,
	ribbon,
	typeInActiveCell,
} from './helpers';

const dialog = (page: Page, name: string) => editor(page).locator(`[data-dialog="${name}"]`);

async function openWhatIf(page: Page, item: string) {
	await ribbon(page).getByRole('tab', { name: 'Data', exact: true }).click();
	await ribbon(page).getByRole('button', { name: 'What-If Analysis' }).click();
	await editor(page).getByRole('menuitem', { name: item }).click();
}

test('Goal Seek solves for an input cell and Cancel restores it', async ({ page }) => {
	await newWorkbook(page);
	await typeInActiveCell(page, '2');
	await typeInActiveCell(page, '=A1*3+4');
	await goToCell(page, 'A2');
	await openWhatIf(page, 'Goal Seek...');
	const box = dialog(page, 'goal-seek');
	await expect(box).toBeVisible();
	await expect(box.getByLabel('Set cell:')).toHaveValue('A2');
	await box.getByLabel('To value:').fill('25');
	await box.getByLabel('By changing cell:').fill('A1');
	await box.getByRole('button', { name: 'OK' }).click();
	const status = dialog(page, 'goal-seek-status');
	await expect(status).toContainText('found a solution');
	await status.getByRole('button', { name: 'OK' }).click();
	await expect(grid(page)).toContainText('25');
	await goToCell(page, 'A1');
	await expect.poll(() => formulaBarValue(page)).toBe('7');

	// A second run that is cancelled leaves the value as it was.
	await goToCell(page, 'A2');
	await openWhatIf(page, 'Goal Seek...');
	await box.getByLabel('To value:').fill('100');
	await box.getByLabel('By changing cell:').fill('A1');
	await box.getByRole('button', { name: 'OK' }).click();
	await status.getByRole('button', { name: 'Cancel' }).click();
	await goToCell(page, 'A1');
	await expect.poll(() => formulaBarValue(page)).toBe('7');
});

test('Data Table fills a one-variable table and protects its cells', async ({ page }) => {
	await newWorkbook(page);
	await typeInActiveCell(page, '3'); // A1: the input cell
	await goToCell(page, 'C1');
	await typeInActiveCell(page, '=A1*10'); // C1: the formula
	await goToCell(page, 'B2');
	await typeInActiveCell(page, '1');
	await typeInActiveCell(page, '2');
	await goToCell(page, 'B1');
	await page.keyboard.press('Shift+ArrowRight');
	await page.keyboard.press('Shift+ArrowDown');
	await page.keyboard.press('Shift+ArrowDown');
	await openWhatIf(page, 'Data Table...');
	const box = dialog(page, 'data-table');
	await expect(box).toBeVisible();
	await box.getByLabel('Column input cell:').fill('A1');
	await box.getByRole('button', { name: 'OK' }).click();
	await expect(box).toBeHidden();
	await goToCell(page, 'C2');
	await expect.poll(() => formulaBarValue(page)).toBe('{=TABLE(,A1)}');
	await expect(grid(page)).toContainText('10');
	await expect(grid(page)).toContainText('20');

	// Typing over part of the table is refused.
	await typeInActiveCell(page, '7');
	await expect(editor(page)).toContainText('Cannot change part of a data table.');
});
