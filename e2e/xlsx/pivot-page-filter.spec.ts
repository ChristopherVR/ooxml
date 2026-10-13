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
		test(`authors ${layout} report filter with columns ${columns}, selection, refresh, history and export`, async ({
			page,
		}) => {
			await newWorkbook(page);
			const rows = [
				['Category', 'Region', 'Amount', 'Quantity', 'Year'],
				['Fruit', 'East', '10', '2', '2025'],
				['Fruit', 'West', '20', '3', '2026'],
				['Veg', 'East', '5', '4', '2025'],
			];
			for (const [r, line] of rows.entries())
				for (const [c, v] of line.entries()) {
					await goToCell(page, `${String.fromCharCode(65 + c)}${r + 1}`);
					await typeInActiveCell(page, v);
				}
			await goToCell(page, 'G1');
			await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
			await ribbon(page).locator('[data-command="insert.pivot-table"]').click();
			const dialog = editor(page).locator('[data-dialog="create-pivot-table"]');
			await dialog.getByLabel('Source Range:', { exact: true }).fill('A1:E4');
			await dialog.getByLabel('Destination Cell:', { exact: true }).fill('G1');
			await dialog.getByLabel('Value Field:', { exact: true }).selectOption('Amount');
			await dialog.getByLabel('Report Layout:', { exact: true }).selectOption(layout);
			await dialog.getByRole('button', { name: 'Add Value Field', exact: true }).click();
			const extra = dialog.locator('[data-pivot-value]');
			await extra.getByLabel('Value Field:', { exact: true }).selectOption('Quantity');
			await extra.getByLabel('Summarize Values By:', { exact: true }).selectOption('count');
			if (columns) await dialog.getByLabel('Column Field:', { exact: true }).selectOption('Year');
			await dialog.getByLabel('Report Filter Field:', { exact: true }).selectOption('Region');
			await dialog
				.getByLabel('Report Filter Item:', { exact: true })
				.selectOption({ label: 'West' });
			await expect(dialog.getByLabel('Second Row Field:', { exact: true })).toBeDisabled();
			await expect(
				dialog.getByLabel('Values Axis:', { exact: true }).locator('option[value=rows]'),
			).toHaveAttribute('disabled');
			await dialog.getByRole('button', { name: 'OK', exact: true }).click();
			const total = columns ? 'J7' : 'H5';
			await goToCell(page, total);
			await expect.poll(() => formulaBarValue(page)).toBe('20');
			await goToCell(page, 'C3');
			await typeInActiveCell(page, '100');
			await goToCell(page, 'G1');
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
						'page-filter.xlsx',
					),
				bytes,
			);
			await goToCell(page, total);
			await expect.poll(() => formulaBarValue(page)).toBe('100');
			await goToCell(page, 'G1');
			await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
			await menu();
			await editor(page)
				.getByRole('menuitem', { name: 'PivotTable Settings', exact: true })
				.click();
			const settings = editor(page).locator('[data-dialog="pivot-settings"]');
			await expect(settings.getByLabel('Report Filter Field:', { exact: true })).toHaveValue(
				'Region',
			);
			await expect(settings.getByLabel('Report Filter Item:', { exact: true })).toHaveValue('1');
			await settings.getByLabel('Report Filter Item:', { exact: true }).selectOption('');
			await settings.getByRole('button', { name: 'OK', exact: true }).click();
			await goToCell(page, 'H1');
			await expect.poll(() => formulaBarValue(page)).toBe('(All)');
			await page.keyboard.press('Control+Z');
			await expect.poll(() => formulaBarValue(page)).toBe('West');
		});
for (const profile of [
	'tabular-all',
	'compact-west',
	'outline-east',
	'tabular-number',
	'tabular-boolean',
	'tabular-blank',
	'tabular-error',
])
	test(`preserves native ${profile} filter/cache definitions and changes selection with history/export`, async ({
		page,
	}) => {
		const bytes = readFileSync(
				new URL(
					`../../src/core/xlsx/__fixtures__/excel-pivot-page-${profile}.xlsx`,
					import.meta.url,
				),
			),
			source = await JSZip.loadAsync(bytes);
		await newWorkbook(page);
		await editor(page).evaluate(
			async (node, data) =>
				(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
					new Uint8Array(data),
					'native-page-filter.xlsx',
				),
			[...bytes],
		);
		const caption =
			profile === 'tabular-all'
				? '(All)'
				: profile === 'compact-west'
					? 'West'
					: profile === 'outline-east'
						? 'East'
						: profile === 'tabular-number'
							? '2025'
							: profile === 'tabular-boolean'
								? "'TRUE"
								: profile === 'tabular-blank'
									? '(blank)'
									: "'#N/A";
		await goToCell(page, 'B1');
		await expect.poll(() => formulaBarValue(page)).toBe(caption);
		const open = async () => {
			await goToCell(page, 'A1');
			await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
			await ribbon(page).getByRole('button', { name: 'PivotTable', exact: true }).last().click();
			await editor(page)
				.getByRole('menuitem', { name: 'PivotTable Settings', exact: true })
				.click();
			return editor(page).locator('[data-dialog="pivot-settings"]');
		};
		let settings = await open();
		await expect(settings.getByLabel('Source Range:', { exact: true })).toBeDisabled();
		await expect(settings.getByLabel('Report Filter Field:', { exact: true })).toHaveValue(
			'Region',
		);
		await settings.getByRole('button', { name: 'OK', exact: true }).click();
		const saved = await editor(page).evaluate(async (node) => [
			...new Uint8Array(
				await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
			),
		]);
		const zip = await JSZip.loadAsync(new Uint8Array(saved));
		for (const part of Object.keys(source.files).filter(
			(n) => n.startsWith('xl/pivot') && !source.files[n]!.dir,
		))
			expect(await zip.file(part)!.async('string')).toBe(await source.file(part)!.async('string'));
		if (profile === 'tabular-all') return;
		settings = await open();
		await settings.getByLabel('Report Filter Item:', { exact: true }).selectOption('');
		await settings.getByRole('button', { name: 'OK', exact: true }).click();
		await goToCell(page, 'B1');
		await expect.poll(() => formulaBarValue(page)).toBe('(All)');
		await page.keyboard.press('Control+Z');
		await expect.poll(() => formulaBarValue(page)).toBe(caption);
		await page.keyboard.press('Control+Y');
		await expect.poll(() => formulaBarValue(page)).toBe('(All)');
		const changed = await editor(page).evaluate(async (node) => [
			...new Uint8Array(
				await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
			),
		]);
		await editor(page).evaluate(
			async (node, data) =>
				(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
					new Uint8Array(data),
					'all-page-filter.xlsx',
				),
			changed,
		);
		await goToCell(page, 'B1');
		await expect.poll(() => formulaBarValue(page)).toBe('(All)');
	});
