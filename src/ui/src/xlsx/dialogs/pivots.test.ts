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

it.each(['tabular', 'compact', 'outline'])(
	'authors %s between Values, preserves settings and converts with history',
	async (form) => {
		const wb = createWorkbook(),
			ctx = createTestContext(wb),
			edit = ctx.session()!;
		[
			['Category', 'Region', 'Amount', 'Quantity'],
			['Fruit', 'East', 10, 2],
			['Fruit', 'West', 20, 3],
		].forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
		ctx.select('A1:D3');
		registerToolDialogs(ctx);
		let done = ctx.dialogs.open('create-pivot-table'),
			dialog = dialogEl(ctx, 'create-pivot-table');
		expect(inputByLabel<HTMLSelectElement>(dialog, 'Values Placement:').disabled).toBe(true);
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Second Row Field:'), 'Region');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Value Field:'), 'Amount');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:'), form);
		clickButton(dialog, 'Add Value Field');
		const extra = dialog.querySelector<HTMLElement>('[data-pivot-value]')!;
		setValue(inputByLabel<HTMLSelectElement>(extra, 'Value Field:'), 'Quantity');
		setValue(inputByLabel<HTMLSelectElement>(extra, 'Summarize Values By:'), 'count');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Values Axis:'), 'rows');
		expect(inputByLabel<HTMLSelectElement>(dialog, 'Values Placement:').disabled).toBe(false);
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Values Placement:'), 'between');
		clickButton(dialog, 'OK');
		await done;
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot).toMatchObject({ layout: form, valuesPosition: 'between', valuesAxis: 'rows' });
		done = ctx.dialogs.open('pivot-settings');
		dialog = dialogEl(ctx, 'pivot-settings');
		expect(inputByLabel<HTMLSelectElement>(dialog, 'Values Placement:').value).toBe('between');
		clickButton(dialog, 'OK');
		await done;
		expect(wb.sheets[0]!.pivotTables![0]).toBe(pivot);
		done = ctx.dialogs.open('pivot-settings');
		dialog = dialogEl(ctx, 'pivot-settings');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Values Axis:'), 'columns');
		expect(inputByLabel<HTMLSelectElement>(dialog, 'Values Placement:').disabled).toBe(true);
		expect(inputByLabel<HTMLSelectElement>(dialog, 'Values Placement:').value).toBe('last');
		clickButton(dialog, 'OK');
		await done;
		expect(wb.sheets[0]!.pivotTables![0]?.valuesAxis).toBe('columns');
		edit.undo();
		expect(wb.sheets[0]!.pivotTables![0]?.valuesPosition).toBe('between');
	},
);

it('keeps compact multiple measures on unchanged settings and switches to outline subtotals with undo', async () => {
	const wb = createWorkbook();
	const ctx = createTestContext(wb);
	const edit = ctx.session()!;
	[
		['Category', 'Region', 'Amount', 'Quantity'],
		['Fruit', 'East', 10, 2],
		['Fruit', 'West', 20, 3],
	].forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
	edit.createPivotTable(
		0,
		{ row: 0, col: 5 },
		{
			source: 'A1:D3',
			rowField: 'Category',
			secondRowField: 'Region',
			valueField: 'Amount',
			layout: 'compact',
			additionalValues: [{ field: 'Quantity', aggregate: 'count' }],
		},
	);
	ctx.select('F1');
	registerToolDialogs(ctx);
	const pivot = wb.sheets[0]!.pivotTables![0]!;
	let done = ctx.dialogs.open('pivot-settings');
	let dialog = dialogEl(ctx, 'pivot-settings');
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:').value).toBe('compact');
	expect(inputByLabel<HTMLInputElement>(dialog, 'Repeat row labels').disabled).toBe(true);
	clickButton(dialog, 'OK');
	await done;
	expect(wb.sheets[0]!.pivotTables![0]).toBe(pivot);
	done = ctx.dialogs.open('pivot-settings');
	dialog = dialogEl(ctx, 'pivot-settings');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:'), 'outline');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Subtotal Position:'), 'bottom');
	expect(inputByLabel<HTMLInputElement>(dialog, 'Repeat row labels').disabled).toBe(false);
	clickButton(dialog, 'OK');
	await done;
	expect(wb.sheets[0]!.pivotTables![0]).toMatchObject({
		layout: 'outline',
		subtotalPosition: 'bottom',
	});
	expect(getCell(wb.sheets[0]!, 4, 5)?.value).toBe('Fruit Total');
	expect(getCell(wb.sheets[0]!, 4, 7)?.value).toBe(30);
	expect(getCell(wb.sheets[0]!, 4, 8)?.value).toBe(2);
	edit.undo();
	expect(wb.sheets[0]!.pivotTables![0]?.layout).toBe('compact');
	expect(getCell(wb.sheets[0]!, 1, 6)?.value).toBe(30);
});

it('authors multiple values in tabular form and preserves no-op settings before removal and undo', async () => {
	const wb = createWorkbook();
	const ctx = createTestContext(wb);
	const edit = ctx.session()!;
	[
		['Category', 'Amount', 'Quantity'],
		['Fruit', 10, 2],
		['Fruit', 20, 3],
		['Veg', 5, 4],
	].forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
	ctx.select('A1:C4');
	registerToolDialogs(ctx);
	let done = ctx.dialogs.open('create-pivot-table');
	let dialog = dialogEl(ctx, 'create-pivot-table');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Value Field:'), 'Amount');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:'), 'tabular');
	clickButton(dialog, 'Add Value Field');
	const extra = dialog.querySelector<HTMLElement>('[data-pivot-value]')!;
	setValue(inputByLabel<HTMLSelectElement>(extra, 'Value Field:'), 'Quantity');
	setValue(inputByLabel<HTMLSelectElement>(extra, 'Summarize Values By:'), 'count');
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:').value).toBe('tabular');
	expect(
		inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:').querySelector<HTMLOptionElement>(
			'[value="compact"]',
		)!.disabled,
	).toBe(false);
	clickButton(dialog, 'OK');
	await done;
	const pivot = wb.sheets[0]!.pivotTables![0]!;
	expect(pivot.additionalValues).toEqual([
		{ field: 'Quantity', aggregate: 'count', caption: 'Count of Quantity' },
	]);
	expect(getCell(wb.sheets[0]!, 1, 5)?.value).toBe(30);
	expect(getCell(wb.sheets[0]!, 1, 6)?.value).toBe(2);
	done = ctx.dialogs.open('pivot-settings');
	dialog = dialogEl(ctx, 'pivot-settings');
	expect(
		inputByLabel<HTMLSelectElement>(dialog.querySelector('[data-pivot-value]')!, 'Value Field:')
			.value,
	).toBe('Quantity');
	clickButton(dialog, 'OK');
	await done;
	expect(wb.sheets[0]!.pivotTables![0]).toBe(pivot);
	done = ctx.dialogs.open('pivot-settings');
	dialog = dialogEl(ctx, 'pivot-settings');
	const measure = dialog.querySelector<HTMLElement>('[data-pivot-value]')!;
	const caption = inputByLabel<HTMLInputElement>(measure, 'Value Caption:');
	setValue(caption, 'Units');
	caption.dispatchEvent(new Event('input', { bubbles: true }));
	setValue(inputByLabel<HTMLSelectElement>(measure, 'Summarize Values By:'), 'max');
	expect(caption.value).toBe('Units');
	clickButton(dialog, 'OK');
	await done;
	expect(wb.sheets[0]!.pivotTables![0]?.additionalValues![0]).toMatchObject({
		caption: 'Units',
		aggregate: 'max',
	});
	done = ctx.dialogs.open('pivot-settings');
	dialog = dialogEl(ctx, 'pivot-settings');
	clickButton(dialog, 'Remove Value Field');
	clickButton(dialog, 'OK');
	await done;
	expect(wb.sheets[0]!.pivotTables![0]?.additionalValues).toBeUndefined();
	edit.undo();
	expect(wb.sheets[0]!.pivotTables![0]?.additionalValues).toHaveLength(1);
});

it('defaults to compact and switches outline subtotal placement with undo', async () => {
	const wb = createWorkbook();
	const ctx = createTestContext(wb);
	const edit = ctx.session()!;
	[
		['Category', 'Region', 'Amount'],
		['Fruit', 'East', 10],
		['Fruit', 'West', 20],
	].forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
	ctx.select('A1:C3');
	registerToolDialogs(ctx);
	let done = ctx.dialogs.open('create-pivot-table');
	let dialog = dialogEl(ctx, 'create-pivot-table');
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:').value).toBe('compact');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Second Row Field:'), 'Region');
	expect(inputByLabel<HTMLInputElement>(dialog, 'Repeat row labels').disabled).toBe(true);
	clickButton(dialog, 'OK');
	await done;
	expect(wb.sheets[0]!.pivotTables![0]).toMatchObject({
		layout: 'compact',
		subtotalPosition: 'top',
	});
	expect(getCell(wb.sheets[0]!, 1, 5)?.value).toBe(30);
	done = ctx.dialogs.open('pivot-settings');
	dialog = dialogEl(ctx, 'pivot-settings');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:'), 'outline');
	expect(inputByLabel<HTMLInputElement>(dialog, 'Repeat row labels').disabled).toBe(false);
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Subtotal Position:'), 'bottom');
	clickButton(dialog, 'OK');
	await done;
	expect(getCell(wb.sheets[0]!, 4, 4)?.value).toBe('Fruit Total');
	expect(getCell(wb.sheets[0]!, 4, 6)?.value).toBe(30);
	edit.undo();
	expect(wb.sheets[0]!.pivotTables![0]?.layout).toBe('compact');
	expect(getCell(wb.sheets[0]!, 1, 5)?.value).toBe(30);
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
	expect(inputByLabel<HTMLInputElement>(dialog, 'Show subtotals').disabled).toBe(true);
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Second Row Field:'), 'Region');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:'), 'tabular');
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
	inputByLabel<HTMLInputElement>(dialog, 'Show subtotals').click();
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

it('authors Values last on rows, preserves unchanged settings and removes the final extra with undo', async () => {
	const wb = createWorkbook();
	const ctx = createTestContext(wb);
	const edit = ctx.session()!;
	[
		['Category', 'Amount', 'Quantity'],
		['Fruit', 10, 2],
		['Fruit', 20, 3],
		['Veg', 5, 4],
	].forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
	ctx.select('A1:C4');
	registerToolDialogs(ctx);
	let done = ctx.dialogs.open('create-pivot-table');
	let dialog = dialogEl(ctx, 'create-pivot-table');
	const axis = inputByLabel<HTMLSelectElement>(dialog, 'Values Axis:');
	expect(axis.disabled).toBe(true);
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Value Field:'), 'Amount');
	clickButton(dialog, 'Add Value Field');
	const extra = dialog.querySelector<HTMLElement>('[data-pivot-value]')!;
	setValue(inputByLabel<HTMLSelectElement>(extra, 'Value Field:'), 'Quantity');
	setValue(inputByLabel<HTMLSelectElement>(extra, 'Summarize Values By:'), 'count');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:'), 'tabular');
	setValue(axis, 'rows');
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:').value).toBe('tabular');
	expect(
		inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:').querySelector<HTMLOptionElement>(
			'[value="compact"]',
		)!.disabled,
	).toBe(false);
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Second Row Field:').disabled).toBe(false);
	const repeat = inputByLabel<HTMLInputElement>(dialog, 'Repeat row labels');
	expect(repeat.disabled).toBe(false);
	repeat.click();
	clickButton(dialog, 'OK');
	await done;
	const pivot = wb.sheets[0]!.pivotTables![0]!;
	expect(pivot).toMatchObject({ valuesAxis: 'rows', layout: 'tabular', repeatLabels: true });
	expect(getCell(wb.sheets[0]!, 2, 4)?.value).toBe('Fruit');
	expect(getCell(wb.sheets[0]!, 2, 6)?.value).toBe(2);
	done = ctx.dialogs.open('pivot-settings');
	dialog = dialogEl(ctx, 'pivot-settings');
	clickButton(dialog, 'OK');
	await done;
	expect(wb.sheets[0]!.pivotTables![0]).toBe(pivot);
	done = ctx.dialogs.open('pivot-settings');
	dialog = dialogEl(ctx, 'pivot-settings');
	clickButton(dialog, 'Remove Value Field');
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Values Axis:').value).toBe('columns');
	clickButton(dialog, 'OK');
	await done;
	expect(wb.sheets[0]!.pivotTables![0]).toMatchObject({ valuesAxis: 'columns' });
	expect(wb.sheets[0]!.pivotTables![0]?.additionalValues).toBeUndefined();
	edit.undo();
	expect(wb.sheets[0]!.pivotTables![0]).toMatchObject({
		valuesAxis: 'rows',
		additionalValues: [{ field: 'Quantity' }],
	});
});

it.each(['compact', 'outline', 'tabular'])(
	'authors two real row fields on the Rows axis in %s layout',
	async (form) => {
		const wb = createWorkbook();
		const ctx = createTestContext(wb);
		const edit = ctx.session()!;
		[
			['Category', 'Region', 'Amount'],
			['Fruit', 'East', 10],
		].forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
		ctx.select('A1:C2');
		registerToolDialogs(ctx);
		const done = ctx.dialogs.open('create-pivot-table');
		const dialog = dialogEl(ctx, 'create-pivot-table');
		clickButton(dialog, 'Add Value Field');
		setValue(
			inputByLabel<HTMLSelectElement>(
				dialog.querySelector<HTMLElement>('[data-pivot-value]')!,
				'Summarize Values By:',
			),
			'count',
		);
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Second Row Field:'), 'Region');
		expect(
			inputByLabel<HTMLSelectElement>(dialog, 'Values Axis:').querySelector<HTMLOptionElement>(
				'[value="rows"]',
			)!.disabled,
		).toBe(false);
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:'), form);
		expect(
			inputByLabel<HTMLSelectElement>(dialog, 'Values Axis:').querySelector<HTMLOptionElement>(
				'[value="rows"]',
			)!.disabled,
		).toBe(false);
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Values Axis:'), 'rows');
		expect(inputByLabel<HTMLSelectElement>(dialog, 'Second Row Field:').disabled).toBe(false);
		const layout = inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:');
		expect(layout.querySelector<HTMLOptionElement>('[value="compact"]')!.disabled).toBe(false);
		expect(layout.querySelector<HTMLOptionElement>('[value="outline"]')!.disabled).toBe(false);
		clickButton(dialog, 'OK');
		await done;
		expect(wb.sheets[0]!.pivotTables![0]).toMatchObject({
			secondRowField: 'Region',
			valuesAxis: 'rows',
			layout: form,
		});
		edit.undo();
		expect(wb.sheets[0]!.pivotTables?.length ?? 0).toBe(0);
	},
);

it('keeps compact Values-row layout and switches to outline repeated labels with history', async () => {
	const wb = createWorkbook();
	const ctx = createTestContext(wb);
	const edit = ctx.session()!;
	[
		['Category', 'Amount', 'Quantity'],
		['Fruit', 10, 2],
		['Veg', 5, 4],
	].forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
	ctx.select('A1:C3');
	registerToolDialogs(ctx);
	let done = ctx.dialogs.open('create-pivot-table');
	let dialog = dialogEl(ctx, 'create-pivot-table');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Value Field:'), 'Amount');
	clickButton(dialog, 'Add Value Field');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Values Axis:'), 'rows');
	expect(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:').value).toBe('compact');
	expect(inputByLabel<HTMLInputElement>(dialog, 'Repeat row labels').disabled).toBe(true);
	clickButton(dialog, 'OK');
	await done;
	expect(wb.sheets[0]!.pivotTables![0]).toMatchObject({ valuesAxis: 'rows', layout: 'compact' });
	expect(getCell(wb.sheets[0]!, 1, 4)?.value).toBe('Fruit');
	expect(getCell(wb.sheets[0]!, 2, 4)?.value).toBe('Sum of Amount');
	done = ctx.dialogs.open('pivot-settings');
	dialog = dialogEl(ctx, 'pivot-settings');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:'), 'outline');
	const repeat = inputByLabel<HTMLInputElement>(dialog, 'Repeat row labels');
	expect(repeat.disabled).toBe(false);
	repeat.click();
	clickButton(dialog, 'OK');
	await done;
	expect(getCell(wb.sheets[0]!, 2, 4)?.value).toBe('Fruit');
	expect(getCell(wb.sheets[0]!, 2, 5)?.value).toBe('Sum of Amount');
	edit.undo();
	expect(wb.sheets[0]!.pivotTables![0]?.layout).toBe('compact');
});

it.each(['compact', 'outline', 'tabular'])(
	'authors first Values with native subtotal eligibility in %s',
	async (layout) => {
		const wb = createWorkbook();
		const ctx = createTestContext(wb);
		const edit = ctx.session()!;
		[
			['Category', 'Region', 'Amount'],
			['Fruit', 'East', 10],
		].forEach((line, r) => line.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
		ctx.select('A1:C2');
		registerToolDialogs(ctx);
		const done = ctx.dialogs.open('create-pivot-table');
		const dialog = dialogEl(ctx, 'create-pivot-table');
		clickButton(dialog, 'Add Value Field');
		setValue(
			inputByLabel<HTMLSelectElement>(
				dialog.querySelector<HTMLElement>('[data-pivot-value]')!,
				'Summarize Values By:',
			),
			'count',
		);
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Second Row Field:'), 'Region');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Report Layout:'), layout);
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Values Axis:'), 'rows');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Values Placement:'), 'first');
		expect(inputByLabel<HTMLSelectElement>(dialog, 'Subtotal Position:').disabled).toBe(
			layout === 'tabular',
		);
		expect(dialog.textContent).toContain('Values are listed before the row fields.');
		clickButton(dialog, 'OK');
		await done;
		expect(wb.sheets[0]!.pivotTables![0]).toMatchObject({ valuesPosition: 'first', layout });
		edit.undo();
		expect(wb.sheets[0]!.pivotTables?.length ?? 0).toBe(0);
	},
);

it.each([
	['de', 'Zeilenbeschriftungen', 'Gesamtergebnis', 'Summe von Amount'],
	['es-MX', 'Etiquetas de fila', 'Total general', 'Suma de Amount'],
	['fr', 'Étiquettes de lignes', 'Total général', 'Somme de Amount'],
	['zh-CN', '行标签', '总计', '求和项:Amount'],
	['en-GB', 'Row Labels', 'Grand Total', 'Sum of Amount'],
])('captions a new report in the editor language %s', async (locale, rowLabels, total, sum) => {
	const wb = createWorkbook();
	const ctx = createTestContext(wb);
	ctx.locale = () => locale;
	const edit = ctx.session()!;
	[
		['Category', 'Amount'],
		['Fruit', 10],
		['Veg', 5],
	].forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
	ctx.select('A1:B3');
	registerToolDialogs(ctx);
	const done = ctx.dialogs.open('create-pivot-table');
	const dialog = dialogEl(ctx, 'create-pivot-table');
	setValue(inputByLabel<HTMLSelectElement>(dialog, 'Value Field:'), 'Amount');
	clickButton(dialog, 'OK');
	await done;
	const pivot = wb.sheets[0]!.pivotTables![0]!;
	expect(pivot.valueCaption).toBe(sum);
	// The default layout is compact: the row header, the value name and the grand total.
	expect(getCell(wb.sheets[0]!, 0, 3)?.value).toBe(rowLabels);
	expect(getCell(wb.sheets[0]!, 0, 4)?.value).toBe(sum);
	expect(getCell(wb.sheets[0]!, 3, 3)?.value).toBe(total);
});
