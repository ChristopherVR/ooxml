import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../workbook';
import { createEditSession } from '../edit/session';
import { getCell } from '../cells';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { pivotGrid } from './aggregate';
import type { PivotAggregate } from '../model';

const setup = () => {
	const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
	const edit = createEditSession(wb);
	const rows = [
		['Category', 'Region', 'Amount'],
		['Fruit', 'East', 10],
		['Fruit', 'West', 20],
		['Veg', 'East', 5],
		['Veg', 'West', 15],
		['Fruit', 'East', 7],
	];
	rows.forEach((row, r) => row.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
	return {
		wb,
		edit,
		create: (aggregate: PivotAggregate = 'sum') =>
			edit.createPivotTable(
				1,
				{ row: 0, col: 0 },
				{
					source: 'Data!A1:C6',
					rowField: 'Category',
					columnField: 'Region',
					valueField: 'Amount',
					aggregate,
				},
			),
	};
};

describe('worksheet-range PivotTables', () => {
	it('matches native Excel numeric aggregation over text-only groups', () => {
		const { wb, create } = setup();
		create();
		const pivot = wb.sheets[1]!.pivotTables![0]!;
		pivot.records = [['Fruit', 'East', 'abc']];
		for (const aggregate of ['sum', 'min', 'max'] as const) {
			pivot.aggregate = aggregate;
			expect(pivotGrid(pivot).values.at(-1)?.at(-1)).toBe(0);
		}
		pivot.aggregate = 'average';
		expect(pivotGrid(pivot).values.at(-1)?.at(-1)).toEqual({ error: '#DIV/0!' });
	});
	it('rejects partial structural edits and restores whole shifts and source renames', () => {
		const { wb, edit, create } = setup();
		create();
		expect(() => edit.insertRows(1, 2, 1)).toThrow('part of a PivotTable');
		expect(getCell(wb.sheets[1]!, 2, 1)?.value).toBe(17);
		expect(getCell(wb.sheets[1]!, 4, 3)?.value).toBe(57);
		edit.insertRows(1, 0, 2);
		expect(wb.sheets[1]!.pivotTables![0]!.location.start.row).toBe(2);
		edit.undo();
		expect(wb.sheets[1]!.pivotTables![0]!.location.start.row).toBe(0);
		edit.renameSheet(0, 'Source Data');
		expect(wb.sheets[1]!.pivotTables![0]!.source).toBe("'Source Data'!A1:C6");
		edit.undo();
		expect(wb.sheets[1]!.pivotTables![0]!.source).toBe('Data!A1:C6');
		expect(() => edit.deleteSheet(0)).toThrow('dependent PivotTables');
	});
	it('copies PivotTables with independent identities and cache output', async () => {
		const { wb, edit, create } = setup();
		create();
		const copy = edit.duplicateSheet(1);
		expect(wb.sheets[copy]!.pivotTables![0]!.name).not.toBe(wb.sheets[1]!.pivotTables![0]!.name);
		const loaded = await loadXlsx(await saveXlsx(wb));
		expect(loaded.sheets[copy]!.pivotTables).toHaveLength(1);
	});
	it('creates a real pivot model with column values and grand totals', () => {
		const { wb, create } = setup();
		create();
		expect(pivotGrid(wb.sheets[1]!.pivotTables![0]!).values).toEqual([
			['Sum of Amount', 'Region', null, null],
			['Category', 'East', 'West', 'Grand Total'],
			['Fruit', 17, 20, 37],
			['Veg', 5, 15, 20],
			['Grand Total', 22, 35, 57],
		]);
		expect(getCell(wb.sheets[1]!, 4, 3)?.value).toBe(57);
	});
	it('refreshes explicitly after input edits and restores cache/output through undo and redo', () => {
		const { wb, edit, create } = setup();
		const name = create();
		edit.setCellInput(0, 1, 2, '100');
		expect(getCell(wb.sheets[1]!, 4, 3)?.value).toBe(57);
		edit.refreshPivotTable(1, name);
		expect(getCell(wb.sheets[1]!, 4, 3)?.value).toBe(147);
		edit.undo();
		expect(getCell(wb.sheets[1]!, 4, 3)?.value).toBe(57);
		edit.redo();
		expect(getCell(wb.sheets[1]!, 4, 3)?.value).toBe(147);
		edit.deletePivotTable(1, name);
		expect(wb.sheets[1]!.pivotTables).toEqual([]);
		expect(getCell(wb.sheets[1]!, 4, 3)).toBeUndefined();
		edit.undo();
		expect(getCell(wb.sheets[1]!, 4, 3)?.value).toBe(147);
	});
	it('computes average grand totals from underlying records and supports count of a row field', () => {
		const { wb, edit, create } = setup();
		const name = create('average');
		expect(getCell(wb.sheets[1]!, 4, 3)?.value).toBe(57 / 5);
		edit.refreshPivotTable(1, name, {
			columnField: '',
			valueField: 'Category',
			aggregate: 'count',
		});
		expect(pivotGrid(wb.sheets[1]!.pivotTables![0]!).values).toEqual([
			['Category', 'Count of Category'],
			['Fruit', 3],
			['Veg', 2],
			['Grand Total', 5],
		]);
		expect(getCell(wb.sheets[1]!, 4, 3)).toBeUndefined();
	});
	it('rejects overlap and direct result editing without corrupting the workbook', () => {
		const { wb, edit, create } = setup();
		create();
		expect(() => create()).toThrow('overlap');
		expect(() => edit.setCellInput(1, 2, 1, '99')).toThrow('cannot be edited');
		expect(() => edit.merge(1, wb.sheets[1]!.pivotTables![0]!.location, 'merge')).toThrow(
			'PivotTable',
		);
		expect(() => edit.createTable(1, wb.sheets[1]!.pivotTables![0]!.location, true)).toThrow(
			'PivotTable',
		);
		expect(getCell(wb.sheets[1]!, 2, 1)?.value).toBe(17);
	});
	it('saves cache records, table definitions and relationships and reloads a refreshable model', async () => {
		const { wb, create } = setup();
		const name = create();
		const once = await loadXlsx(await saveXlsx(wb));
		expect(once.warnings).toEqual([]);
		expect(once.sheets[1]!.pivotTables![0]).toMatchObject({
			name,
			aggregate: 'sum',
			source: 'Data!A1:C6',
			rowField: 'Category',
			columnField: 'Region',
			valueField: 'Amount',
		});
		const second = await loadXlsx(await saveXlsx(once));
		expect(second.sheets[1]!.pivotTables![0]?.records).toEqual(
			wb.sheets[1]!.pivotTables![0]?.records,
		);
		const edit = createEditSession(second);
		edit.setCellInput(0, 1, 2, '100');
		edit.refreshPivotTable(1, name);
		const third = await loadXlsx(await saveXlsx(second));
		expect(getCell(third.sheets[1]!, 4, 3)?.value).toBe(147);
		expect(third.sheets[1]!.pivotTables).toHaveLength(1);
	});
});
