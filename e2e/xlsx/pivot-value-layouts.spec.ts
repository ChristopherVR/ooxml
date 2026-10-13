import { expect, test } from '@playwright/test';
import { resolve } from 'node:path';
import { editor, formulaBarValue, goToCell, ribbon } from './helpers';

test('switches native multiple measures across compact, outline and tabular layouts, undoes and reopens', async ({
	page,
}, info) => {
	await page.goto('/');
	await page
		.locator('#landing-file')
		.setInputFiles(
			resolve(
				process.cwd(),
				'../../src/core/xlsx/__fixtures__/excel-pivot-values-compact-rows-top-columns.xlsx',
			),
		);
	await goToCell(page, 'H10');
	await expect.poll(() => formulaBarValue(page)).toBe('57');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	const settings = editor(page).locator('[data-dialog="pivot-settings"]');
	const open = async () => {
		await ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
		await editor(page).getByRole('menuitem', { name: 'PivotTable Settings', exact: true }).click();
	};
	await open();
	await expect(settings.getByLabel('Report Layout:', { exact: true })).toHaveValue('compact');
	await expect(settings.locator('[data-pivot-value]')).toHaveCount(2);
	await settings.getByLabel('Report Layout:', { exact: true }).selectOption('outline');
	await settings.getByLabel('Subtotal Position:', { exact: true }).selectOption('bottom');
	await settings.getByLabel('Repeat row labels', { exact: true }).check();
	await page.screenshot({ path: info.outputPath('multiple-outline-settings.png') });
	await settings.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'I12');
	await expect.poll(() => formulaBarValue(page)).toBe('57');
	await goToCell(page, 'A7');
	await expect.poll(() => formulaBarValue(page)).toBe('Fruit Total');
	await page.keyboard.press('Control+Z');
	await goToCell(page, 'H10');
	await expect.poll(() => formulaBarValue(page)).toBe('57');
	await page.keyboard.press('Control+Y');
	await goToCell(page, 'I12');
	await expect.poll(() => formulaBarValue(page)).toBe('57');
	await open();
	await settings.getByLabel('Report Layout:', { exact: true }).selectOption('tabular');
	await settings.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'I10');
	await expect.poll(() => formulaBarValue(page)).toBe('57');
	await open();
	await settings.getByLabel('Report Layout:', { exact: true }).selectOption('compact');
	await settings.getByLabel('Subtotal Position:', { exact: true }).selectOption('top');
	await expect(settings.getByLabel('Repeat row labels', { exact: true })).toBeDisabled();
	await settings.getByRole('button', { name: 'OK', exact: true }).click();
	await goToCell(page, 'H10');
	await expect.poll(() => formulaBarValue(page)).toBe('57');
	const bytes = await editor(page).evaluate(async (node) => [
		...new Uint8Array(
			await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
		),
	]);
	await editor(page).evaluate(
		async (node, data) =>
			(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'measure-layouts.xlsx',
			),
		bytes,
	);
	await goToCell(page, 'J10');
	await expect.poll(() => formulaBarValue(page)).toBe('11.4');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await open();
	await expect(settings.getByLabel('Report Layout:', { exact: true })).toHaveValue('compact');
	await expect(settings.locator('[data-pivot-value]')).toHaveCount(2);
	await expect(settings.getByLabel('Repeat row labels', { exact: true })).toBeChecked();
	await settings.getByRole('button', { name: 'Cancel', exact: true }).click();
});
