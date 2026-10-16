import { expect, it, afterEach } from 'vitest';
import { createWorkbook, getCell, loadXlsx, saveXlsx } from 'ooxml-core/xlsx';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import {
	createTestContext,
	dialogEl,
	inputByLabel,
	setValue,
	clickButton,
} from '../commands/test-support';
import { registerToolDialogs } from './register-tools';
afterEach(() => {
	document.body.innerHTML = '';
});
it.each(['compact', 'outline', 'tabular'])(
	'authors %s report selection and changes All/remove with history',
	async (layout) => {
		const wb = createWorkbook(),
			ctx = createTestContext(wb),
			edit = ctx.session()!;
		[
			['Category', 'Region', 'Amount', 'Year'],
			['Fruit', 'East', 10, 2025],
			['Fruit', 'West', 20, 2026],
			['Veg', 'East', 5, 2025],
		].forEach((line, r) => line.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
		ctx.select('A1:D4');
		registerToolDialogs(ctx);
		let done = ctx.dialogs.open('create-pivot-table'),
			dialog = dialogEl(ctx, 'create-pivot-table');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Value Field:'), 'Amount');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:'), layout);
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Column Field:'), 'Year');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Field:'), 'Region');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Item:'), '1');
		expect(inputByLabel<HTMLSelectElement>(dialog, 'Second Row Field:').disabled).toBe(false);
		expect(
			inputByLabel<HTMLSelectElement>(dialog, 'Second Row Field:').querySelector<HTMLOptionElement>(
				'option[value=Region]',
			)!.disabled,
		).toBe(true);
		expect(
			inputByLabel<HTMLSelectElement>(dialog, 'Values Axis:').querySelector<HTMLOptionElement>(
				'[value=rows]',
			)!.disabled,
		).toBe(false);
		clickButton(dialog, 'OK');
		await done;
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.pageFilter).toEqual({ field: 'Region', item: 'West' });
		expect(getCell(wb.sheets[0]!, 0, 6)?.value).toBe('West');
		done = ctx.dialogs.open('pivot-settings');
		dialog = dialogEl(ctx, 'pivot-settings');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Item:'), '');
		clickButton(dialog, 'OK');
		await done;
		expect(wb.sheets[0]!.pivotTables![0]!.pageFilter).toEqual({ field: 'Region' });
		edit.undo();
		expect(wb.sheets[0]!.pivotTables![0]!.pageFilter?.item).toBe('West');
		edit.redo();
		done = ctx.dialogs.open('pivot-settings');
		dialog = dialogEl(ctx, 'pivot-settings');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Field:'), '');
		expect(inputByLabel<HTMLSelectElement>(dialog, 'Second Row Field:').disabled).toBe(false);
		clickButton(dialog, 'OK');
		await done;
		expect(wb.sheets[0]!.pivotTables![0]!.pageFilter).toBeUndefined();
	},
);
it.each(['tabular', 'compact', 'outline'])(
	'authors %s nested filters without dropping the selected page field',
	async (layout) => {
		const wb = createWorkbook(),
			ctx = createTestContext(wb),
			edit = ctx.session()!;
		[
			['Category', 'Region', 'Market', 'Amount', 'Year', 'Quantity'],
			['Fruit', 'North', 'East', 10, 2025, 2],
			['Fruit', 'South', 'West', 20, 2026, 3],
			['Veg', 'North', 'East', 5, 2025, 4],
		].forEach((line, r) => line.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
		ctx.select('A1:F4');
		registerToolDialogs(ctx);
		let done = ctx.dialogs.open('create-pivot-table'),
			dialog = dialogEl(ctx, 'create-pivot-table');
		setValue(inputByLabel<HTMLInputElement>(dialog, 'Destination Cell:'), 'H1');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Value Field:'), 'Amount');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Field:'), 'Market');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Item:'), '1');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Second Row Field:'), 'Region');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:'), layout);
		expect(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Field:').value).toBe('Market');
		expect(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Item:').value).toBe('1');
		clickButton(dialog, 'OK');
		await done;
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.secondRowField).toBe('Region');
		expect(pivot.pageFilter).toEqual({ field: 'Market', item: 'West' });
		expect(pivot.records).toHaveLength(3);
		ctx.select('H1');
		done = ctx.dialogs.open('pivot-settings');
		dialog = dialogEl(ctx, 'pivot-settings');
		clickButton(dialog, 'OK');
		await done;
		expect(wb.sheets[0]!.pivotTables![0]).toBe(pivot);
		done = ctx.dialogs.open('pivot-settings');
		dialog = dialogEl(ctx, 'pivot-settings');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Item:'), '');
		clickButton(dialog, 'OK');
		await done;
		expect(wb.sheets[0]!.pivotTables![0]!.pageFilter).toEqual({ field: 'Market' });
		edit.undo();
		expect(wb.sheets[0]!.pivotTables![0]!.pageFilter).toEqual({ field: 'Market', item: 'West' });
		expect(wb.sheets[0]!.pivotTables![0]!.secondRowField).toBe('Region');
	},
);
it('preserves untouched native filter/cache definitions when accepting settings', async () => {
	const bytes = readFileSync('../core/xlsx/__fixtures__/excel-pivot-page-compact-west.xlsx');
	const wb = await loadXlsx(bytes),
		ctx = createTestContext(wb);
	ctx.select('A1');
	registerToolDialogs(ctx);
	const done = ctx.dialogs.open('pivot-settings'),
		dialog = dialogEl(ctx, 'pivot-settings');
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Field:').value).toBe('Region');
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Item:').value).toBe('1');
	clickButton(dialog, 'OK');
	await done;
	const source = await JSZip.loadAsync(bytes),
		saved = await JSZip.loadAsync(await saveXlsx(wb));
	for (const name of Object.keys(source.files).filter(
		(n) => n.startsWith('xl/pivot') && !source.files[n]!.dir,
	))
		expect(await saved.file(name)!.async('string')).toBe(await source.file(name)!.async('string'));
});

it('retains a selected report field when adding a default value field', async () => {
	const wb = createWorkbook(),
		ctx = createTestContext(wb),
		edit = ctx.session()!;
	[
		['Category', 'Region', 'Amount', 'Quantity', 'Year'],
		['Fruit', 'East', 10, 2, 2025],
	].forEach((line, r) => line.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
	ctx.select('A1:E2');
	registerToolDialogs(ctx);
	const done = ctx.dialogs.open('create-pivot-table'),
		dialog = dialogEl(ctx, 'create-pivot-table');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Value Field:'), 'Amount');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Field:'), 'Year');
	clickButton(dialog, 'Add Value Field');
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Field:').value).toBe('Year');
	expect(
		inputByLabel<HTMLSelectElement>(
			dialog.querySelector<HTMLElement>('[data-pivot-value]')!,
			'Value Field:',
		).value,
	).toBe('Quantity');
	clickButton(dialog, 'Cancel');
	await done;
});

it('compares unchanged public filter properties independently of object key order', async () => {
	const wb = createWorkbook(),
		ctx = createTestContext(wb),
		edit = ctx.session()!;
	[
		['Category', 'Region', 'Amount'],
		['Fruit', 'East', 10],
		['Veg', 'West', 20],
	].forEach((line, r) => line.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
	edit.createPivotTable(
		0,
		{ row: 0, col: 6 },
		{
			source: 'A1:C3',
			rowField: 'Category',
			valueField: 'Amount',
			pageFilter: { item: 'West', field: 'Region' },
		},
	);
	const pivot = wb.sheets[0]!.pivotTables![0]!;
	ctx.select('G1');
	registerToolDialogs(ctx);
	const done = ctx.dialogs.open('pivot-settings'),
		dialog = dialogEl(ctx, 'pivot-settings');
	clickButton(dialog, 'OK');
	await done;
	expect(wb.sheets[0]!.pivotTables![0]).toBe(pivot);
});
it('authors a report filter above Values on rows and keeps it when the axis changes', async () => {
	const wb = createWorkbook(),
		ctx = createTestContext(wb),
		edit = ctx.session()!;
	[
		['Category', 'Region', 'Amount', 'Quantity'],
		['Fruit', 'East', 10, 2],
		['Fruit', 'West', 20, 3],
		['Veg', 'East', 5, 4],
	].forEach((line, r) => line.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
	ctx.select('A1:D4');
	registerToolDialogs(ctx);
	let done = ctx.dialogs.open('create-pivot-table'),
		dialog = dialogEl(ctx, 'create-pivot-table');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Value Field:'), 'Amount');
	clickButton(dialog, 'Add Value Field');
	const extra = dialog.querySelector<HTMLElement>('[data-pivot-value]')!;
	setValue(inputByLabel<HTMLSelectElement>(extra, 'Value Field:'), 'Quantity');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Values Axis:'), 'rows');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Field:'), 'Region');
	// The filter field stays selected and enabled with Values on rows.
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Field:').value).toBe('Region');
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Field:').disabled).toBe(false);
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Item:'), '1');
	clickButton(dialog, 'OK');
	await done;
	let pivot = wb.sheets[0]!.pivotTables![0]!;
	expect(pivot).toMatchObject({
		valuesAxis: 'rows',
		pageFilter: { field: 'Region', item: 'West' },
	});
	expect(getCell(wb.sheets[0]!, 0, 6)?.value).toBe('West');
	done = ctx.dialogs.open('pivot-settings');
	dialog = dialogEl(ctx, 'pivot-settings');
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Values Axis:').value).toBe('rows');
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Report Filter Field:').value).toBe('Region');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Values Axis:'), 'columns');
	clickButton(dialog, 'OK');
	await done;
	pivot = wb.sheets[0]!.pivotTables![0]!;
	expect(pivot.valuesAxis).toBe('columns');
	expect(pivot.pageFilter).toEqual({ field: 'Region', item: 'West' });
	edit.undo();
	expect(wb.sheets[0]!.pivotTables![0]!.valuesAxis).toBe('rows');
});
