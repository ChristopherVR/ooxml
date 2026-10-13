import { expect, test } from '@playwright/test';
import { editor, goToCell, grid, newWorkbook, ribbon, typeInActiveCell } from './helpers';

test('authors, saves and clears a sparkline date axis with keyboard undo', async ({
	page,
}, info) => {
	await newWorkbook(page);
	for (const [ref, value] of [
		['A1', '10'],
		['B1', '20'],
		['C1', '30'],
		['A2', '46000'],
		['B2', '46001'],
		['C2', '46020'],
	]) {
		await goToCell(page, ref!);
		await typeInActiveCell(page, value!);
	}
	await goToCell(page, 'F1');
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await ribbon(page).locator('[data-command="insert.sparkline-line"]').click();
	const create = editor(page).locator('[data-dialog="create-sparklines"]');
	await create.getByLabel('Data Range:', { exact: true }).fill('A1:C1');
	await create.getByLabel('Date Range (optional):', { exact: true }).fill('A2:C2');
	await create.getByRole('button', { name: 'OK', exact: true }).click();
	const line = grid(page).locator('.xg-c:not([hidden]) .xg-spark polyline');
	await expect(line).toHaveCount(1);
	const ratio = async () => {
		const points = (await line.getAttribute('points'))!
			.split(' ')
			.map((pair) => Number(pair.split(',')[0]));
		return (points[1]! - points[0]!) / (points[2]! - points[0]!);
	};
	expect(await ratio()).toBeCloseTo(0.05, 2);
	const bytes = await editor(page).evaluate(async (node) => [
		...new Uint8Array(
			await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
		),
	]);
	await editor(page).evaluate(
		async (node, data) =>
			(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'date-axis.xlsx',
			),
		bytes,
	);
	await goToCell(page, 'F1');
	expect(await ratio()).toBeCloseTo(0.05, 2);
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	await ribbon(page).locator('[data-command="sparkline.settings"]').first().click();
	const settings = editor(page).locator('[data-dialog="sparkline-settings"]');
	await expect(settings.getByLabel('Date Range (optional):', { exact: true })).toHaveValue(
		"'Sheet1'!A2:C2",
	);
	await settings.getByLabel('Date Range (optional):', { exact: true }).fill('');
	await page.screenshot({ path: info.outputPath('sparkline-date-axis.png') });
	await settings.getByRole('button', { name: 'OK', exact: true }).click();
	expect(await ratio()).toBeCloseTo(0.5, 2);
	await page.keyboard.press('Control+Z');
	expect(await ratio()).toBeCloseTo(0.05, 2);
});
