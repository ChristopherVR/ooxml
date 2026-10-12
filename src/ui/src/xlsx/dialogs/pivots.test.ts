import { afterEach, expect, it } from 'vitest';
import { createWorkbook, getCell } from 'ooxml-core/xlsx';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	setValue,
} from '../commands/test-support';
import { registerToolDialogs } from './register-tools';
import { pivotCommands } from '../commands/pivots';

afterEach(() => {
	document.body.innerHTML = '';
});

it('authors hierarchical row fields and reconfigures subtotals with one undo step', async () => {
	const wb = createWorkbook();
	const ctx = createTestContext(wb);
	const edit = ctx.session()!;
	[
		['Category', 'Region', 'Amount'],
		['Fruit', 'East', 10],
		['Fruit', 'West', 20],
		['Veg', 'West', 5],
	].forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
	ctx.select('A1:C4');
	registerToolDialogs(ctx);
	let done = ctx.dialogs.open('create-pivot-table');
	let dialog = dialogEl(ctx, 'create-pivot-table');
	expect(inputByLabel<HTMLInputElement>(dialog, 'Show subtotals below each group').disabled).toBe(
		true,
	);
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Second Row Field:'), 'Region');
	clickButton(dialog, 'OK');
	expect(await done).toBe(true);
	expect(wb.sheets[0]!.pivotTables![0]).toMatchObject({
		rowField: 'Category',
		secondRowField: 'Region',
		subtotals: true,
	});
	expect(getCell(wb.sheets[0]!, 3, 4)?.value).toBe('Fruit Total');
	done = ctx.dialogs.open('pivot-settings');
	dialog = dialogEl(ctx, 'pivot-settings');
	inputByLabel<HTMLInputElement>(dialog, 'Show subtotals below each group').click();
	inputByLabel<HTMLInputElement>(dialog, 'Repeat row labels').click();
	clickButton(dialog, 'OK');
	expect(await done).toBe(true);
	expect(wb.sheets[0]!.pivotTables![0]).toMatchObject({ subtotals: false, repeatLabels: true });
	expect(getCell(wb.sheets[0]!, 2, 4)?.value).toBe('Fruit');
	edit.undo();
	expect(getCell(wb.sheets[0]!, 3, 4)?.value).toBe('Fruit Total');
});

it('authors fields from source headers and leaves unchanged settings alone', async () => {
	const wb = createWorkbook();
	const ctx = createTestContext(wb);
	const edit = ctx.session()!;
	[
		['Category', 'Amount'],
		['Fruit', 10],
		['Fruit', 20],
	].forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
	ctx.select('A1:B3');
	registerToolDialogs(ctx);
	const done = ctx.dialogs.open('create-pivot-table');
	const dialog = dialogEl(ctx, 'create-pivot-table');
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Value Field:').value).toBe('Amount');
	clickButton(dialog, 'OK');
	expect(await done).toBe(true);
	const pivot = wb.sheets[0]!.pivotTables![0]!;
	pivot.valueCaption = 'Custom caption';
	const settingsDone = ctx.dialogs.open('pivot-settings');
	const settings = dialogEl(ctx, 'pivot-settings');
	clickButton(settings, 'OK');
	expect(await settingsDone).toBe(true);
	expect(wb.sheets[0]!.pivotTables![0]).toBe(pivot);
	expect(pivot.valueCaption).toBe('Custom caption');
	const changeDone = ctx.dialogs.open('pivot-settings');
	const change = dialogEl(ctx, 'pivot-settings');
	setValue(inputByLabel<HTMLSelectElement>(change, 'Summarize Values By:'), 'count');
	clickButton(change, 'OK');
	await changeDone;
	expect(getCell(wb.sheets[0]!, 2, 4)?.value).toBe(2);
	ctx.commands.registerAll(pivotCommands());
	ctx.commands.run('pivot.delete');
	expect(wb.sheets[0]!.pivotTables).toEqual([]);
	edit.undo();
	expect(wb.sheets[0]!.pivotTables).toHaveLength(1);
});

it('guards authoring commands when the sheet is protected', () => {
	const wb = createWorkbook();
	const ctx = createTestContext(wb);
	wb.sheets[0]!.protection = { sheet: true };
	for (const command of pivotCommands()) expect(command.enabled?.(ctx)).toBe(false);
});
