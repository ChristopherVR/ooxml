// @vitest-environment jsdom
import { createWorkbook, getCell } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	pressKey,
	setValue,
} from '../commands/test-support.js';
import { seriesValues } from './fill-series.js';
import { registerToolDialogs } from './register-tools.js';
import { splitLine } from './text-to-columns.js';

afterEach(() => (document.body.innerHTML = ''));

function setup() {
	const ctx = createTestContext(createWorkbook());
	registerToolDialogs(ctx);
	return ctx;
}
const ws = (ctx: ReturnType<typeof setup>) => ctx.workbook()!.sheets[0]!;
const v = (ctx: ReturnType<typeof setup>, row: number, col: number) =>
	getCell(ws(ctx), row, col)?.value;

describe('Protect Sheet', () => {
	it('protects with the chosen allowed actions; Cancel does nothing', async () => {
		const ctx = setup();
		const cancelled = ctx.dialogs.open('protect-sheet');
		clickButton(dialogEl(ctx, 'protect-sheet'), 'Cancel');
		await cancelled;
		expect(ws(ctx).protection).toBeUndefined();
		const result = ctx.dialogs.open('protect-sheet');
		const dialog = dialogEl(ctx, 'protect-sheet');
		setValue(inputByLabel(dialog, 'Password to unprotect sheet:'), 'a');
		setValue(inputByLabel(dialog, 'Reenter password to proceed:'), 'b');
		clickButton(dialog, 'OK');
		expect(ctx.toasts.at(-1)?.message).toBe('Confirmation password is not identical.');
		setValue(inputByLabel(dialog, 'Password to unprotect sheet:'), '');
		setValue(inputByLabel(dialog, 'Reenter password to proceed:'), '');
		inputByLabel(dialog, 'Format cells').click();
		clickButton(dialog, 'OK');
		await result;
		expect(ws(ctx).protection).toEqual({
			sheet: true,
			allow: ['selectLockedCells', 'selectUnlockedCells', 'formatCells'],
		});
	});
});

describe('Insert / Delete cells', () => {
	it('shifts cells down and deletes an entire row', async () => {
		const ctx = setup();
		ctx.session()!.setCellInput(0, 0, 0, 'x');
		ctx.select('A1');
		const ins = ctx.dialogs.open('insert-cells');
		inputByLabel(dialogEl(ctx, 'insert-cells'), 'Shift cells down').click();
		clickButton(dialogEl(ctx, 'insert-cells'), 'OK');
		expect(await ins).toBe('down');
		expect(v(ctx, 1, 0)).toBe('x');
		ctx.select('A1');
		const del = ctx.dialogs.open('delete-cells');
		inputByLabel(dialogEl(ctx, 'delete-cells'), 'Entire row').click();
		clickButton(dialogEl(ctx, 'delete-cells'), 'OK');
		await del;
		expect(v(ctx, 0, 0)).toBe('x');
		const esc = ctx.dialogs.open('delete-cells');
		pressKey(dialogEl(ctx, 'delete-cells'), 'Escape');
		expect(await esc).toBeUndefined();
		expect(v(ctx, 0, 0)).toBe('x');
	});
});

describe('Remove Duplicates', () => {
	it('removes duplicate rows under a header', async () => {
		const ctx = setup();
		ctx.session()!.setRangeValues(0, { row: 0, col: 0 }, [
			['Name', 'N'],
			['a', 1],
			['b', 2],
			['A', 1],
			['c', 3],
		]);
		ctx.select('A1');
		const result = ctx.dialogs.open('remove-duplicates');
		const dialog = dialogEl(ctx, 'remove-duplicates');
		expect(inputByLabel(dialog, 'My data has headers').checked).toBe(true);
		expect(inputByLabel(dialog, 'Name').checked).toBe(true);
		clickButton(dialog, 'OK');
		expect(await result).toBe(1);
		expect([1, 2, 3, 4].map((r) => v(ctx, r, 0))).toEqual(['a', 'b', 'c', undefined]);
		expect(ctx.toasts[0]?.message).toBe(
			'{removed} duplicate values found and removed; {kept} unique values remain.'
				.replace('{removed}', '1')
				.replace('{kept}', '3'),
		);
	});
});

describe('Text to Columns', () => {
	it('splits with qualifiers and consecutive delimiters', () => {
		expect(splitLine('"a,b",c', { delimiters: [','], consecutive: false, qualifier: '"' })).toEqual(
			['a,b', 'c'],
		);
		expect(splitLine('a  b', { delimiters: [' '], consecutive: true, qualifier: '' })).toEqual([
			'a',
			'b',
		]);
	});

	it('splits the column at commas into the destination', async () => {
		const ctx = setup();
		ctx.session()!.setRangeValues(0, { row: 0, col: 0 }, [['x,1'], ['y,2']]);
		ctx.select('A1:A2');
		const result = ctx.dialogs.open('text-to-columns');
		const dialog = dialogEl(ctx, 'text-to-columns');
		inputByLabel(dialog, 'Comma').click();
		expect(dialog.querySelectorAll('.xve-results td').length).toBe(4);
		clickButton(dialog, 'OK');
		await result;
		expect([v(ctx, 0, 0), v(ctx, 0, 1), v(ctx, 1, 0), v(ctx, 1, 1)]).toEqual(['x', 1, 'y', 2]);
		ctx.session()!.undo();
		expect(v(ctx, 0, 0)).toBe('x,1');
	});
});

describe('Symbol', () => {
	it('inserts the chosen symbol into the active cell', async () => {
		const ctx = setup();
		ctx.session()!.setCellInput(0, 0, 0, 'Price ');
		const result = ctx.dialogs.open('symbol');
		const dialog = dialogEl(ctx, 'symbol');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Subset:'), 'currency');
		dialog.querySelector<HTMLButtonElement>('[aria-label="U+20AC"]')!.click();
		expect(dialog.textContent).toContain('U+20AC');
		clickButton(dialog, 'Insert');
		expect(await result).toBe('€');
		expect(v(ctx, 0, 0)).toBe('Price €');
	});
});

describe('Series', () => {
	it('computes linear, growth and month series', () => {
		expect(seriesValues(1, 4, 'linear', 2)).toEqual([1, 3, 5, 7]);
		expect(seriesValues(1, 5, 'growth', 2, { stop: 10 })).toEqual([1, 2, 4, 8]);
		expect(seriesValues(45306, 2, 'date', 1, { unit: 'month' })).toEqual([45306, 45337]);
	});

	it('fills the selected row with a step', async () => {
		const ctx = setup();
		ctx.session()!.setCellInput(0, 0, 0, '5');
		ctx.select('A1:D1');
		const result = ctx.dialogs.open('fill-series');
		const dialog = dialogEl(ctx, 'fill-series');
		setValue(inputByLabel(dialog, 'Step value:'), '5');
		clickButton(dialog, 'OK');
		await result;
		expect([0, 1, 2, 3].map((c) => v(ctx, 0, c))).toEqual([5, 10, 15, 20]);
	});
});

describe('Create Table', () => {
	it('creates a table over the current region with headers', async () => {
		const ctx = setup();
		ctx.session()!.setRangeValues(0, { row: 0, col: 0 }, [
			['Item', 'Qty'],
			['a', 1],
		]);
		ctx.select('A1');
		const result = ctx.dialogs.open('create-table', { styleName: 'TableStyleLight9' });
		const dialog = dialogEl(ctx, 'create-table');
		expect(inputByLabel(dialog, 'Where is the data for your table?').value).toBe('=$A$1:$B$2');
		clickButton(dialog, 'OK');
		await result;
		expect(ws(ctx).tables[0]).toMatchObject({ styleName: 'TableStyleLight9', headerRow: true });
	});

	it('keeps the dialog open when the core refuses an overlap', async () => {
		const ctx = setup();
		ctx.session()!.setRangeValues(0, { row: 0, col: 0 }, [['Item'], ['a']]);
		ctx.session()!.createTable(0, { start: { row: 0, col: 0 }, end: { row: 1, col: 0 } }, true);
		ctx.select('A1');
		void ctx.dialogs.open('create-table');
		clickButton(dialogEl(ctx, 'create-table'), 'OK');
		await Promise.resolve();
		expect(ctx.toasts[0]?.kind).toBe('error');
		expect(ws(ctx).tables).toHaveLength(1);
		expect(dialogEl(ctx, 'create-table')).toBeTruthy();
	});
});
