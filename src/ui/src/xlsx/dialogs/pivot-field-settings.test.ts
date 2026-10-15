import { afterEach, describe, expect, it } from 'vitest';
import {
	createWorkbook,
	pivotGrid,
	type PivotTable,
	type PivotTableOptions,
} from 'ooxml-core/xlsx';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	setValue,
} from '../commands/test-support';
import { pivotCommands } from '../commands/pivots';
import { registerToolDialogs } from './register-tools';

afterEach(() => {
	document.body.innerHTML = '';
});

const setup = (extra: Partial<PivotTableOptions> = {}) => {
	const wb = createWorkbook();
	const ctx = createTestContext(wb);
	const edit = ctx.session()!;
	[
		['Category', 'Region', 'Year', 'Amount'],
		['Fruit', 'East', '2025', 10],
		['Fruit', 'West', '2026', 20],
		['Veg', 'East', '2025', 5],
		['Veg', 'West', '2026', 15],
	].forEach((line, r) => line.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
	edit.createPivotTable(
		0,
		{ row: 0, col: 6 },
		{
			source: 'A1:D5',
			rowField: 'Region',
			columnField: 'Category',
			moreColumnFields: ['Year'],
			valueField: 'Amount',
			layout: 'compact',
			...extra,
		},
	);
	ctx.select('G1');
	registerToolDialogs(ctx);
	const pivot = (): PivotTable => wb.sheets[0]!.pivotTables![0]!;
	return { ctx, edit, pivot };
};
const command = (id: string) => pivotCommands().find((candidate) => candidate.id === id)!;
const box = (dialog: HTMLElement, label: string) =>
	dialog.querySelector<HTMLInputElement>(`[aria-label="${label}"]`)!;
const toggle = (input: HTMLInputElement, checked: boolean) => {
	input.checked = checked;
	input.dispatchEvent(new Event('change', { bubbles: true }));
};

describe('Field Settings for column fields', () => {
	it('is available once the columns nest two fields', () => {
		expect(command('pivot.field-settings').enabled?.(setup().ctx)).toBe(true);
		const single = setup({ moreColumnFields: [] });
		expect(command('pivot.field-settings').enabled?.(single.ctx)).toBe(false);
	});

	it('offers the column fields, sets custom subtotals and lays out their columns', async () => {
		const { ctx, edit, pivot } = setup();
		const done = ctx.dialogs.open('pivot-field-settings');
		const dialog = dialogEl(ctx, 'pivot-field-settings');
		const chosen = inputByLabel<HTMLSelectElement>(dialog, 'Field:');
		expect([...chosen.options].map((option) => option.value)).toEqual(['Category', 'Year']);
		toggle(dialog.querySelector<HTMLInputElement>('input[type=radio][value=custom]')!, true);
		toggle(box(dialog, 'Sum'), true);
		toggle(box(dialog, 'Average'), true);
		clickButton(dialog, 'OK');
		await done;
		expect(pivot().fieldSubtotals).toEqual([{ field: 'Category', functions: ['sum', 'average'] }]);
		expect(pivotGrid(pivot()).values[1]).toEqual([
			null,
			'Fruit',
			null,
			'Fruit Sum',
			'Fruit Average',
			'Veg',
			null,
			'Veg Sum',
			'Veg Average',
			'Grand Total',
		]);
		edit.undo();
		expect(pivot().fieldSubtotals).toBeUndefined();
	});

	it('reloads the saved choice and removes the subtotal columns for None', async () => {
		const { ctx, pivot } = setup({
			fieldSubtotals: [{ field: 'Category', functions: ['max'] }],
		});
		const done = ctx.dialogs.open('pivot-field-settings');
		const dialog = dialogEl(ctx, 'pivot-field-settings');
		expect(box(dialog, 'Max').checked).toBe(true);
		toggle(dialog.querySelector<HTMLInputElement>('input[type=radio][value=none]')!, true);
		clickButton(dialog, 'OK');
		await done;
		expect(pivot().fieldSubtotals).toEqual([{ field: 'Category', functions: [] }]);
		expect(pivotGrid(pivot()).values[1]).toEqual([null, 'Fruit', null, 'Veg', null, 'Grand Total']);
	});

	it('keeps the setting of a field when another field is changed', async () => {
		const { ctx, pivot } = setup({
			fieldSubtotals: [{ field: 'Category', functions: ['max'] }],
		});
		const done = ctx.dialogs.open('pivot-field-settings');
		const dialog = dialogEl(ctx, 'pivot-field-settings');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Field:'), 'Year');
		toggle(dialog.querySelector<HTMLInputElement>('input[type=radio][value=none]')!, true);
		clickButton(dialog, 'OK');
		await done;
		expect(pivot().fieldSubtotals).toEqual([
			{ field: 'Category', functions: ['max'] },
			{ field: 'Year', functions: [] },
		]);
	});
});
