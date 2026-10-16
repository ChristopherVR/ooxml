import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import {
	editor,
	newWorkbook,
	ribbon,
	goToCell,
	typeInActiveCell,
	formulaBarValue,
} from './helpers';

for (const layout of ['tabular', 'compact', 'outline'])
	for (const columns of [false, true])
		test(`authors nested ${layout} report filter with columns ${columns}, history and reopen`, async ({
			page,
		}) => {
			await newWorkbook(page);
			const rows = [
				['Category', 'Region', 'Market', 'Year', 'Amount', 'Quantity'],
				['Fruit', 'North', 'East', '2025', '10', '2'],
				['Fruit', 'South', 'West', '2026', '20', '3'],
				['Veg', 'North', 'East', '2025', '5', '4'],
			];
			for (const [r, row] of rows.entries())
				for (const [c, value] of row.entries()) {
					await goToCell(page, `${String.fromCharCode(65 + c)}${r + 1}`);
					await typeInActiveCell(page, value);
				}
			await goToCell(page, 'H1');
			await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
			await ribbon(page).locator('[data-command="insert.pivot-table"]').click();
			const dialog = editor(page).locator('[data-dialog="create-pivot-table"]');
			await dialog.getByLabel('Source Range:', { exact: true }).fill('A1:F4');
			await dialog.getByLabel('Destination Cell:', { exact: true }).fill('H1');
			await dialog.getByLabel('Value Field:', { exact: true }).selectOption('Amount');
			await dialog.getByLabel('Second Row Field:', { exact: true }).selectOption('Region');
			await dialog.getByLabel('Report Layout:', { exact: true }).selectOption(layout);
			if (layout !== 'tabular')
				await dialog.getByLabel('Subtotal Position:', { exact: true }).selectOption('bottom');
			await dialog.getByRole('button', { name: 'Add Value Field', exact: true }).click();
			await dialog
				.locator('[data-pivot-value]')
				.getByLabel('Value Field:', { exact: true })
				.selectOption('Quantity');
			if (columns) await dialog.getByLabel('Column Field:', { exact: true }).selectOption('Year');
			await expect(dialog.getByLabel('Report Filter Field:', { exact: true })).toBeEnabled();
			await dialog.getByLabel('Report Filter Field:', { exact: true }).selectOption('Market');
			await dialog
				.getByLabel('Report Filter Item:', { exact: true })
				.selectOption({ label: 'West' });
			await expect(dialog.getByLabel('Second Row Field:', { exact: true })).toBeEnabled();
			await expect(
				dialog.getByLabel('Second Row Field:', { exact: true }).locator('option[value=Market]'),
			).toHaveAttribute('disabled');
			await expect(
				dialog.getByLabel('Values Axis:', { exact: true }).locator('option[value=rows]'),
			).not.toHaveAttribute('disabled');
			await dialog.getByRole('button', { name: 'OK', exact: true }).click();
			const total = `${layout === 'compact' ? 'I' : 'J'}${(layout === 'tabular' ? 6 : 7) + (columns ? 2 : 0)}`;
			await goToCell(page, total);
			await expect.poll(() => formulaBarValue(page)).toBe('20');
			await goToCell(page, 'E3');
			await typeInActiveCell(page, '100');
			await goToCell(page, 'H1');
			const menu = () =>
				ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
			await menu();
			await editor(page).getByRole('menuitem', { name: 'Refresh PivotTable', exact: true }).click();
			await goToCell(page, total);
			await expect.poll(() => formulaBarValue(page)).toBe('100');
			await page.keyboard.press('Control+Z');
			await expect.poll(() => formulaBarValue(page)).toBe('20');
			await page.keyboard.press('Control+Y');
			await expect.poll(() => formulaBarValue(page)).toBe('100');
			const bytes = await editor(page).evaluate(async (node) => [
				...new Uint8Array(
					await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
				),
			]);
			await editor(page).evaluate(
				async (node, data) =>
					(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
						new Uint8Array(data),
						'nested-filter.xlsx',
					),
				bytes,
			);
			await goToCell(page, total);
			await expect.poll(() => formulaBarValue(page)).toBe('100');
			await goToCell(page, 'H1');
			await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
			await menu();
			await editor(page)
				.getByRole('menuitem', { name: 'PivotTable Settings', exact: true })
				.click();
			const settings = editor(page).locator('[data-dialog="pivot-settings"]');
			await expect(settings.getByLabel('Second Row Field:', { exact: true })).toHaveValue('Region');
			await expect(settings.getByLabel('Report Filter Field:', { exact: true })).toHaveValue(
				'Market',
			);
			await settings.getByLabel('Report Filter Item:', { exact: true }).selectOption('');
			await settings.getByRole('button', { name: 'OK', exact: true }).click();
			await goToCell(page, 'I1');
			await expect.poll(() => formulaBarValue(page)).toBe('(All)');
			await page.keyboard.press('Control+Z');
			await expect.poll(() => formulaBarValue(page)).toBe('West');
		});

for (const profile of [
	'tabular-east-bottom-columns',
	'compact-west-top-columns',
	'outline-east-none-columns',
	'tabular-east-bottom-columns-repeat',
	'outline-west-none-columns-repeat',
])
	test(`unchanged native nested filter settings preserve ${profile}`, async ({ page }) => {
		await newWorkbook(page);
		const source = readFileSync(
			new URL(
				`../../src/core/xlsx/__fixtures__/excel-pivot-nested-page-${profile}.xlsx`,
				import.meta.url,
			),
		);
		await editor(page).evaluate(
			async (node, bytes) =>
				(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
					new Uint8Array(bytes),
					'native-nested-filter.xlsx',
				),
			[...source],
		);
		await goToCell(page, 'A1');
		await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
		await ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
		await editor(page).getByRole('menuitem', { name: 'PivotTable Settings', exact: true }).click();
		const dialog = editor(page).locator('[data-dialog="pivot-settings"]');
		await expect(dialog.getByLabel('Second Row Field:', { exact: true })).toHaveValue('Region');
		await expect(dialog.getByLabel('Report Filter Field:', { exact: true })).toHaveValue('Market');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		const saved = await editor(page).evaluate(async (node) => [
			...new Uint8Array(
				await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
			),
		]);
		const original = await JSZip.loadAsync(source),
			zip = await JSZip.loadAsync(new Uint8Array(saved));
		for (const part of Object.keys(original.files).filter(
			(p) => p.startsWith('xl/pivot') && !original.files[p]!.dir,
		))
			expect(await zip.file(part)!.async('string')).toBe(
				await original.file(part)!.async('string'),
			);
		await editor(page).evaluate(
			async (node, bytes) =>
				(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
					new Uint8Array(bytes),
					'reopened.xlsx',
				),
			saved,
		);
		await goToCell(page, 'B1');
		await expect
			.poll(() => formulaBarValue(page))
			.toBe(profile.includes('-west-') ? 'West' : 'East');
	});
