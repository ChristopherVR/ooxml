import { expect, test } from '@playwright/test';
import { editor, goToCell, newWorkbook, ribbon, typeInActiveCell } from './helpers';

test('inserts a recommendation, charts the existing report and uses field buttons from the keyboard', async ({
	page,
}) => {
	await newWorkbook(page);
	for (const [address, value] of [
		['A1', 'Category'],
		['B1', 'Amount'],
		['A2', 'Fruit'],
		['B2', '30'],
		['A3', 'Vegetables'],
		['B3', '10'],
	] as const) {
		await goToCell(page, address);
		await typeInActiveCell(page, value);
	}
	await goToCell(page, 'A1');
	const root = editor(page);
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await ribbon(page).getByRole('button', { name: 'Recommended PivotTables', exact: true }).click();
	const recommend = root.locator('[data-dialog="recommended-pivot-tables"]');
	await expect(recommend.getByRole('table', { name: 'Preview' })).toContainText('Fruit');
	await recommend.getByRole('button', { name: 'Insert', exact: true }).click();
	const charts = ribbon(page).getByRole('button', { name: 'Charts', exact: true });
	if (await charts.isVisible()) await charts.click();
	await root.getByRole('button', { name: 'PivotChart', exact: true }).click();
	const chart = root.locator('.xg-chart');
	const button = chart.getByRole('button', { name: 'Axis field: Category' });
	await expect(button).toBeVisible();
	await button.focus();
	await button.press('Enter');
	const filter = root.locator('[data-dialog="pivot-field-filter"]');
	await filter.getByRole('checkbox', { name: 'Fruit', exact: true }).uncheck();
	await filter.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(chart.locator('.xg-obj-desc')).toContainText('Vegetables');
	await expect(chart.locator('.xg-obj-desc')).not.toContainText('Fruit');
	await ribbon(page).getByRole('tab', { name: 'Chart Design', exact: true }).click();
	await ribbon(page).getByRole('button', { name: 'Field Buttons', exact: true }).click();
	await expect(button).toHaveCount(0);
});
