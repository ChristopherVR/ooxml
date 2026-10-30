import { expect, test, type Page } from '@playwright/test';
import { editor, goToCell, newWorkbook, ribbon, typeInActiveCell } from './helpers';

const dialog = (page: Page, name: string) => editor(page).locator(`[data-dialog="${name}"]`);
const arrows = (page: Page) => editor(page).locator('[data-part="audit-arrows"] line');

async function formulasTab(page: Page) {
	await ribbon(page).getByRole('tab', { name: 'Formulas', exact: true }).click();
}

async function seed(page: Page) {
	await newWorkbook(page);
	await typeInActiveCell(page, '2'); // A1
	await typeInActiveCell(page, '3'); // A2
	await goToCell(page, 'B1');
	await typeInActiveCell(page, '=A1+A2');
}

test('Trace Precedents draws arrows and Remove Arrows clears them', async ({ page }) => {
	await seed(page);
	await goToCell(page, 'B1');
	await formulasTab(page);
	await ribbon(page).getByRole('button', { name: 'Trace Precedents' }).click();

	await expect(arrows(page)).toHaveCount(2);
	await ribbon(page).getByRole('button', { name: 'Remove Arrows' }).first().click();
	await expect(arrows(page)).toHaveCount(0);
});

test('Evaluate Formula steps once', async ({ page }) => {
	await seed(page);
	await goToCell(page, 'B1');
	await formulasTab(page);
	await ribbon(page).getByRole('button', { name: 'Evaluate Formula' }).click();
	const box = dialog(page, 'evaluate-formula');
	await expect(box).toBeVisible();
	await expect(box.locator('u')).toHaveText('A1');
	await box.getByRole('button', { name: 'Evaluate' }).click();
	await expect(box).toContainText('=2+A2');
	await box.locator('button.xve-btn', { hasText: 'Close' }).click();
	await expect(box).toBeHidden();
});

test('Watch Window lists an added watch', async ({ page }) => {
	await seed(page);
	await goToCell(page, 'B1');
	await formulasTab(page);
	await ribbon(page).getByRole('button', { name: 'Watch Window' }).click();
	const pane = editor(page).locator('[data-part="watch-window"]');
	await expect(pane).toBeVisible();
	await pane.getByRole('button', { name: 'Add Watch...' }).click();
	const add = dialog(page, 'add-watch');
	await expect(add).toBeVisible();
	await add.getByRole('button', { name: 'Add', exact: true }).click();
	await expect(pane).toContainText('=A1+A2');
	await expect(pane).toContainText('5');
	await expect(pane).toContainText('Book1');
});
