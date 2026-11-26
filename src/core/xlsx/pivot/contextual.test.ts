import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { createEditSession } from '../edit/session';
import { getCell } from '../cells';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { pivotGrid } from './aggregate';
import type { PivotTablePatch } from '../model';

const fixture = (name: string) =>
	loadXlsx(readFileSync(new URL(`../__fixtures__/excel-pivot${name}.xlsx`, import.meta.url)));

async function report(suffix = '-rows') {
	const workbook = await fixture(suffix);
	const edit = createEditSession(workbook);
	const ws = workbook.sheets[0]!;
	const read = () => ws.pivotTables![0]!;
	const patch = (value: PivotTablePatch) => edit.refreshPivotTable(0, read().name, value);
	const reopen = async () => (await loadXlsx(await saveXlsx(workbook))).sheets[0]!.pivotTables![0]!;
	return { workbook, edit, ws, read, patch, reopen };
}

it('rebuilds native layouts, headers, blank rows, totals and styles in one undoable edit', async () => {
	const { edit, patch, read, reopen } = await report();
	const before = pivotGrid(read()).values;
	patch({
		layout: 'outline',
		repeatLabels: true,
		subtotalPosition: 'top',
		subtotals: false,
		fieldSubtotals: [{ field: 'Region', functions: [] }],
		display: {
			colGrandTotals: false,
			showHeaders: false,
			showDrill: false,
			blankRows: true,
			style: { name: 'PivotStyleMedium4', showRowStripes: true, showColStripes: true },
		},
	});
	const values = pivotGrid(read()).values;
	expect(values[0]).toEqual([null, null, 'Sum of Amount']);
	expect(values).toContainEqual([null, null, null]);
	expect(values.flat()).not.toContain('Grand Total');
	expect(values.flat()).not.toContain('Fruit Total');
	const reopened = await reopen();
	expect(reopened).toMatchObject({
		layout: 'outline',
		repeatLabels: true,
		display: read().display,
	});
	expect(pivotGrid(reopened).values).toEqual(values);
	expect(reopened.sourceXml).toContain('insertBlankRow="1"');
	expect(reopened.sourceXml).toContain('showRowStripes="1"');
	edit.undo();
	expect(pivotGrid(read()).values).toEqual(before);
	edit.redo();
	expect(pivotGrid(read()).values).toEqual(values);
});

it('collapses and expands outer labels, and applies row and column grand totals independently', async () => {
	const { patch, reopen, read } = await report();
	patch({ display: { collapsed: true } });
	expect(pivotGrid(await reopen()).values).toEqual([
		['Category', 'Region', 'Sum of Amount'],
		['Fruit', null, 37],
		['Veg', null, 20],
		['Grand Total', null, 57],
	]);
	patch({ display: { collapsed: false } });
	expect(pivotGrid(read()).values).toHaveLength(8);
	const cols = await report('');
	cols.patch({ display: { rowGrandTotals: false } });
	expect(pivotGrid(await cols.reopen()).values.at(-1)).toEqual(['Grand Total', 22, 35]);
	cols.patch({ display: { rowGrandTotals: true, colGrandTotals: false } });
	expect(pivotGrid(await cols.reopen()).values.at(-1)).toEqual(['Veg', 5, 15, 20]);
});

it('edits caption and value conditions without filtering the backing cache', async () => {
	const { edit, patch, reopen, read } = await report();
	patch({ condition: { kind: 'captionEqual', value: 'Veg' } });
	let reopened = await reopen();
	expect(reopened.condition).toEqual({ kind: 'captionEqual', value: 'Veg' });
	expect(pivotGrid(reopened).values.at(-1)).toEqual(['Grand Total', null, 20]);
	expect(reopened.records).toHaveLength(5);
	patch({ condition: { kind: 'valueGreaterThan', measure: 0, value: 20 } });
	reopened = await reopen();
	expect(reopened.condition).toEqual({ kind: 'valueGreaterThan', measure: 0, value: 20 });
	expect(pivotGrid(reopened).values.at(-1)).toEqual(['Grand Total', null, 37]);
	edit.undo();
	expect(read().condition?.kind).toBe('captionEqual');
	patch({ condition: null });
	expect(pivotGrid(await reopen()).values.at(-1)).toEqual(['Grand Total', null, 57]);
});

it('evaluates a calculated measure from aggregate source values and round trips the formula', async () => {
	const { edit, patch, read, reopen } = await report();
	patch({ calculatedField: { name: 'DoubleAmount', field: 'Amount', factor: 2 } });
	const calculated = await reopen();
	expect(calculated.calculatedField).toEqual({
		name: 'DoubleAmount',
		field: 'Amount',
		factor: 2,
		caption: 'Sum of DoubleAmount',
	});
	expect(pivotGrid(calculated).values.at(-1)).toEqual(['Grand Total', null, 57, 114]);
	edit.undo();
	expect(read().calculatedField).toBeUndefined();
	edit.redo();
	patch({ calculatedField: null });
	expect(pivotGrid(await reopen()).values.at(-1)).toEqual(['Grand Total', null, 57]);
});

it('renames, moves and redirects a report and rolls back Refresh All if a report fails', async () => {
	const { workbook, edit, ws, patch, read, reopen } = await report();
	const original = read().name;
	patch({ name: 'NewSales', destination: { row: 3, col: 5 }, source: 'Data!A1:C5' });
	expect(read().location.start).toEqual({ row: 3, col: 5 });
	expect(getCell(ws, 0, 0)?.value ?? null).toBeNull();
	expect(await reopen()).toMatchObject({
		name: 'NewSales',
		source: expect.stringContaining('A1:C5'),
	});
	edit.undo();
	expect(read().name).toBe(original);
	const other = edit.createPivotTable(
		0,
		{ row: 12, col: 0 },
		{
			source: 'Data!A1:C6',
			rowField: 'Category',
			valueField: 'Amount',
		},
	);
	edit.setCellInput(1, 1, 2, '30');
	edit.refreshAllPivotTables();
	expect(pivotGrid(read()).values.at(-1)?.at(-1)).toBe(77);
	expect(pivotGrid(ws.pivotTables![1]!).values.at(-1)?.at(-1)).toBe(77);
	edit.undo();
	expect(pivotGrid(read()).values.at(-1)?.at(-1)).toBe(57);
	expect(pivotGrid(ws.pivotTables![1]!).values.at(-1)?.at(-1)).toBe(57);
	const failed = ws.pivotTables!.find((pivot) => pivot.name === other)!;
	failed.source = 'Missing!A1:C6';
	expect(() => edit.refreshAllPivotTables()).toThrow();
	expect(getCell(workbook.sheets[0]!, 7, 2)?.value).toBe(57);
});
