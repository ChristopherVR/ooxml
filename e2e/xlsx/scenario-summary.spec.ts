import { expect, test, type Page } from '@playwright/test';
import { openFixture } from './fixture-files';
import {
	clickRibbonControl,
	editor,
	FRAMEWORKS,
	grid,
	pageErrors,
	ribbon,
	sheetTabs,
} from './helpers';

// openpyxl-whatif.xlsx: sheet "WhatIf" with A1=3, A2=4, A3=A1*A2 and the scenarios Low (1, 2) and
// High (30, 40) on A1:A2.
const dialog = (page: Page, name: string) => editor(page).locator(`[data-dialog="${name}"]`);

async function openScenarioManager(page: Page) {
	await ribbon(page).getByRole('tab', { name: 'Data', exact: true }).click();
	await clickRibbonControl(
		page,
		ribbon(page).getByRole('button', { name: 'What-If Analysis', exact: true }),
	);
	await editor(page).getByRole('menuitem', { name: 'Scenario Manager...' }).click();
	await expect(dialog(page, 'scenario-manager')).toBeVisible();
}

for (const framework of FRAMEWORKS)
	test(`${framework}: Scenario Summary builds the report and Merge copies scenarios`, async ({
		page,
	}) => {
		const errors = pageErrors(page);
		await openFixture(page, framework, 'openpyxl-whatif.xlsx');
		await openScenarioManager(page);
		const manager = dialog(page, 'scenario-manager');
		await expect(manager.getByRole('option')).toHaveCount(2);
		await manager.getByRole('button', { name: 'Summary...', exact: true }).click();
		const summary = dialog(page, 'scenario-summary');
		await summary.getByLabel('Result cells:').fill('A3');
		await summary.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(summary).toBeHidden();
		await manager.locator('button.xve-btn', { hasText: 'Close' }).click();

		// The report is a new, active sheet that names both scenarios and their results.
		await expect(sheetTabs(page).getByRole('tab')).toHaveCount(2);
		await expect(sheetTabs(page).getByRole('tab', { selected: true })).toHaveText(
			/Scenario Summary/,
		);
		await expect(grid(page)).toContainText('Low');
		await expect(grid(page)).toContainText('High');
		await expect(grid(page)).toContainText('Changing Cells');
		await expect(grid(page)).toContainText('1200');

		// Merge: a sheet with no scenarios of its own takes both from the first sheet.
		await sheetTabs(page).getByRole('button', { name: 'New sheet' }).click();
		await openScenarioManager(page);
		await expect(manager.getByRole('option')).toHaveCount(0);
		await manager.getByRole('button', { name: 'Merge...', exact: true }).click();
		const merge = dialog(page, 'scenario-merge');
		await merge.getByLabel('Merge scenarios from sheet:').selectOption({ label: 'WhatIf' });
		await merge.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(merge).toBeHidden();
		await expect(manager.getByRole('option')).toHaveCount(2);
		await expect(manager.getByRole('option').first()).toContainText('Low');
		expect(errors).toEqual([]);
	});
