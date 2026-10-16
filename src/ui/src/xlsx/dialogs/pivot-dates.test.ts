import { afterEach, describe, expect, it } from 'vitest';
import { createWorkbook, pivotGrid, type PivotTable } from 'ooxml-core/xlsx';
import { EDITOR_LOCALES, STRINGS } from 'ooxml-core/xlsx/ui';
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

const setup = (extra: object = {}) => {
	const wb = createWorkbook();
	const ctx = createTestContext(wb);
	const edit = ctx.session()!;
	// 45292 is 2024-01-01 in the 1900 date system.
	[
		['Date', 'Region', 'Product', 'Amount'],
		[45292, 'East', 'Apple', 10],
		[45330, 'West', 'Apple', 20],
		[45413, 'East', 'Pear', 5],
		[45657, 'West', 'Pear', 15],
	].forEach((line, r) => line.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
	edit.applyStyle(0, [{ start: { row: 1, col: 0 }, end: { row: 4, col: 0 } }], {
		numFmt: 'yyyy-mm-dd',
	});
	edit.createPivotTable(
		0,
		{ row: 0, col: 6 },
		{
			source: 'A1:D5',
			rowField: 'Date',
			valueField: 'Amount',
			layout: 'compact',
			...extra,
		},
	);
	ctx.select('G1');
	registerToolDialogs(ctx);
	const pivot = (): PivotTable => wb.sheets[0]!.pivotTables![0]!;
	return { wb, ctx, edit, pivot };
};
const command = (id: string) => pivotCommands().find((candidate) => candidate.id === id)!;
const box = (dialog: HTMLElement, label: string, index = 0) =>
	dialog.querySelectorAll<HTMLInputElement>(`[aria-label="${label}"]`)[index]!;
const toggle = (input: HTMLInputElement, checked: boolean) => {
	input.checked = checked;
	input.dispatchEvent(new Event('change', { bubbles: true }));
};

describe('Group Selection on a date field', () => {
	it('groups by months and years, replaces the date on the rows and undoes', async () => {
		const { ctx, edit, pivot } = setup();
		expect(command('pivot.group-selection').enabled?.(ctx)).toBe(true);
		const done = ctx.dialogs.open('pivot-group-dates');
		const dialog = dialogEl(ctx, 'pivot-group-dates');
		expect(inputByLabel<HTMLSelectElement>(dialog, 'Field:').value).toBe('Date');
		expect(box(dialog, 'Months').checked).toBe(true);
		expect(inputByLabel(dialog, 'Starting at:').value).toBe('2024-01-01');
		expect(inputByLabel(dialog, 'Ending at:').value).toBe('2024-12-31');
		expect(inputByLabel(dialog, 'Number of days:').disabled).toBe(true);
		toggle(box(dialog, 'Years'), true);
		clickButton(dialog, 'OK');
		await done;
		expect(pivot().dateGroups).toEqual([{ field: 'Date', by: ['months', 'years'] }]);
		expect(pivot()).toMatchObject({ rowField: 'Years (Date)', secondRowField: 'Months (Date)' });
		expect(pivotGrid(pivot()).values.map((line) => line[0])).toEqual([
			'Row Labels',
			'2024',
			'Jan',
			'Feb',
			'May',
			'Dec',
			'Grand Total',
		]);
		edit.undo();
		expect(pivot().rowField).toBe('Date');
		expect(pivot().dateGroups).toBeUndefined();
	});

	it('groups the dates of a report that has a report filter', async () => {
		const { ctx, pivot } = setup({ pageFilter: { field: 'Product', item: 'Pear' } });
		const done = ctx.dialogs.open('pivot-group-dates');
		const dialog = dialogEl(ctx, 'pivot-group-dates');
		toggle(box(dialog, 'Years'), true);
		clickButton(dialog, 'OK');
		await done;
		expect(pivot().pageFilter).toEqual({ field: 'Product', item: 'Pear' });
		expect(pivot().dateGroups).toEqual([{ field: 'Date', by: ['months', 'years'] }]);
		// The group range still comes from all the dates, so the pages of only Pear stay in place.
		expect(pivotGrid(pivot()).values.map((line) => line[0])).toEqual([
			'Product',
			null,
			'Row Labels',
			'2024',
			'May',
			'Dec',
			'Grand Total',
		]);
	});

	it('takes a chosen range and a number of days when grouping by days alone', async () => {
		const { ctx, pivot } = setup();
		const done = ctx.dialogs.open('pivot-group-dates');
		const dialog = dialogEl(ctx, 'pivot-group-dates');
		toggle(box(dialog, 'Months'), false);
		toggle(box(dialog, 'Days'), true);
		expect(inputByLabel(dialog, 'Number of days:').disabled).toBe(false);
		setValue(inputByLabel(dialog, 'Number of days:'), '100');
		toggle(box(dialog, 'Auto', 0), false);
		setValue(inputByLabel(dialog, 'Starting at:'), '2023-12-31');
		toggle(box(dialog, 'Auto', 1), false);
		setValue(inputByLabel(dialog, 'Ending at:'), '2024-12-31');
		clickButton(dialog, 'OK');
		await done;
		const [start, end] = [45291, 45657];
		expect(pivot().dateGroups).toEqual([{ field: 'Date', by: ['days'], days: 100, start, end }]);
		expect(pivotGrid(pivot()).values.map((line) => line[0])).toEqual([
			'Row Labels',
			'2023-12-31 - 2024-04-08',
			'2024-04-09 - 2024-07-17',
			'2024-10-26 - 2024-12-31',
			'Grand Total',
		]);
	});

	it('reloads an existing group, offers its fields on the rows and ungroups', async () => {
		const { ctx, edit, pivot } = setup({
			dateGroups: [{ field: 'Date', by: ['quarters', 'years'] }],
		});
		const done = ctx.dialogs.open('pivot-group-dates');
		const group = dialogEl(ctx, 'pivot-group-dates');
		expect(box(group, 'Quarters').checked).toBe(true);
		expect(box(group, 'Months').checked).toBe(false);
		clickButton(group, 'Cancel');
		await done;
		const settings = ctx.dialogs.open('pivot-settings');
		const dialog = dialogEl(ctx, 'pivot-settings');
		const row = inputByLabel<HTMLSelectElement>(dialog, 'Row Field:');
		expect(row.value).toBe('Years (Date)');
		expect([...row.options].map((option) => option.value)).toEqual([
			'Region',
			'Product',
			'Amount',
			'Quarters (Date)',
			'Years (Date)',
		]);
		clickButton(dialog, 'Cancel');
		await settings;
		expect(command('pivot.ungroup').enabled?.(ctx)).toBe(true);
		command('pivot.ungroup').run(ctx);
		expect(pivot().rowField).toBe('Date');
		edit.undo();
		expect(pivot().rowField).toBe('Years (Date)');
	});

	it('rejects a date that cannot be read and an empty selection of units', async () => {
		const { ctx, pivot } = setup();
		const done = ctx.dialogs.open('pivot-group-dates');
		const dialog = dialogEl(ctx, 'pivot-group-dates');
		toggle(box(dialog, 'Auto', 0), false);
		setValue(inputByLabel(dialog, 'Starting at:'), 'last week');
		clickButton(dialog, 'OK');
		expect(ctx.toasts.at(-1)).toMatchObject({ message: 'Enter a valid date.', kind: 'error' });
		toggle(box(dialog, 'Auto', 0), true);
		toggle(box(dialog, 'Months'), false);
		clickButton(dialog, 'OK');
		expect(ctx.toasts.at(-1)).toMatchObject({ message: 'Choose at least one group.' });
		expect(pivot().dateGroups).toBeUndefined();
		clickButton(dialog, 'Cancel');
		await done;
	});
});

describe('many row fields and field settings', () => {
	const rows = () => {
		const state = setup({ rowField: 'Region', secondRowField: 'Product' });
		return state;
	};

	it('adds and removes row fields in the PivotTable settings', async () => {
		const { ctx, pivot } = rows();
		let done = ctx.dialogs.open('pivot-settings');
		let dialog = dialogEl(ctx, 'pivot-settings');
		clickButton(dialog, 'Add Row Field');
		const extra = dialog.querySelector<HTMLElement>('[data-pivot-row]')!;
		setValue(inputByLabel<HTMLSelectElement>(extra, 'Row Field:'), 'Amount');
		clickButton(dialog, 'OK');
		await done;
		expect(pivot().moreRowFields).toEqual(['Amount']);
		expect(pivotGrid(pivot()).values.map((line) => line[0])).toContain('Pear');
		done = ctx.dialogs.open('pivot-settings');
		dialog = dialogEl(ctx, 'pivot-settings');
		expect(
			inputByLabel<HTMLSelectElement>(dialog, 'Values Axis:').querySelector<HTMLOptionElement>(
				'[value=rows]',
			)!.disabled,
		).toBe(false);
		clickButton(dialog, 'Remove Row Field');
		clickButton(dialog, 'OK');
		await done;
		expect(pivot().moreRowFields).toBeUndefined();
	});

	it('adds and removes column fields in the PivotTable settings', async () => {
		const { ctx, pivot } = setup({ columnField: 'Product' });
		let done = ctx.dialogs.open('pivot-settings');
		let dialog = dialogEl(ctx, 'pivot-settings');
		clickButton(dialog, 'Add Column Field');
		const extra = dialog.querySelector<HTMLElement>('[data-pivot-column]')!;
		setValue(inputByLabel<HTMLSelectElement>(extra, 'Column Field:'), 'Region');
		clickButton(dialog, 'OK');
		await done;
		expect(pivot()).toMatchObject({ columnField: 'Product', moreColumnFields: ['Region'] });
		expect(pivotGrid(pivot()).values[1]).toContain('Apple Total');
		done = ctx.dialogs.open('pivot-settings');
		dialog = dialogEl(ctx, 'pivot-settings');
		expect(dialog.querySelector('[data-pivot-column]')).not.toBeNull();
		clickButton(dialog, 'Remove Column Field');
		clickButton(dialog, 'OK');
		await done;
		expect(pivot().moreColumnFields).toBeUndefined();
	});

	it('sets Automatic, None and custom subtotals per row field', async () => {
		const { ctx, edit, pivot } = rows();
		expect(command('pivot.field-settings').enabled?.(ctx)).toBe(true);
		let done = ctx.dialogs.open('pivot-field-settings');
		let dialog = dialogEl(ctx, 'pivot-field-settings');
		expect(box(dialog, 'Sum').disabled).toBe(true);
		toggle(dialog.querySelector<HTMLInputElement>('input[type=radio][value=custom]')!, true);
		expect(box(dialog, 'Sum').disabled).toBe(false);
		clickButton(dialog, 'OK');
		expect(ctx.toasts.at(-1)).toMatchObject({ message: 'Choose at least one function.' });
		toggle(box(dialog, 'Average'), true);
		toggle(box(dialog, 'Max'), true);
		toggle(box(dialog, 'StdDev'), true);
		clickButton(dialog, 'OK');
		await done;
		expect(pivot().fieldSubtotals).toEqual([
			{ field: 'Region', functions: ['average', 'max', 'stdDev'] },
		]);
		expect(
			pivotGrid(pivot())
				.values.map((line) => line[0])
				.filter((label) => String(label).startsWith('East ')),
		).toEqual(['East Average', 'East Max', 'East StdDev']);
		done = ctx.dialogs.open('pivot-field-settings');
		dialog = dialogEl(ctx, 'pivot-field-settings');
		expect(box(dialog, 'Average').checked).toBe(true);
		toggle(dialog.querySelector<HTMLInputElement>('input[type=radio][value=none]')!, true);
		clickButton(dialog, 'OK');
		await done;
		expect(pivot().fieldSubtotals).toEqual([{ field: 'Region', functions: [] }]);
		expect(
			pivotGrid(pivot()).values.some(
				(line) => String(line[0]).endsWith(' Total') && line[0] !== 'Grand Total',
			),
		).toBe(false);
		done = ctx.dialogs.open('pivot-field-settings');
		dialog = dialogEl(ctx, 'pivot-field-settings');
		toggle(dialog.querySelector<HTMLInputElement>('input[type=radio][value=auto]')!, true);
		clickButton(dialog, 'OK');
		await done;
		expect(pivot().fieldSubtotals).toBeUndefined();
		edit.undo();
		expect(pivot().fieldSubtotals).toEqual([{ field: 'Region', functions: [] }]);
	});

	it('is unavailable for a single row field', () => {
		const { ctx } = setup();
		expect(command('pivot.field-settings').enabled?.(ctx)).toBe(false);
	});
});

describe('translations of the new text', () => {
	const KEYS = [
		'Group Selection',
		'Field Settings',
		'Auto',
		'Seconds',
		'Minutes',
		'Hours',
		'Number of days:',
		'Add Row Field',
		'Remove Row Field',
		'Additional Row Field',
		'Add Column Field',
		'Remove Column Field',
		'Additional Column Field',
		'Choose valid PivotTable column fields.',
		'After {field}',
		'Values are listed after {field}.',
		'Report filters support one column field.',
		'Subtotals',
		'Select one or more functions:',
		'Product',
		'Count Numbers',
		'StdDev',
		'StdDevp',
		'Var',
		'Varp',
		'Grouping',
		'Enter a valid date.',
		'Choose at least one group.',
		'Choose at least one function.',
		'Choose a valid date group.',
		'Only fields that contain dates can be grouped by date.',
		'The date group must end after it starts.',
		'A source field already has the name of a date group field.',
		'The date group has too many items.',
		'A date field with groups is replaced on the axes by its group fields.',
		'A date field grouped by several units can only be used as a row field.',
		'Choose valid PivotTable row fields.',
		'Choose valid field subtotals.',
	];
	it.each(EDITOR_LOCALES)('%s has every new key', (locale) => {
		expect(KEYS.filter((key) => !STRINGS[locale][key])).toEqual([]);
	});
});
