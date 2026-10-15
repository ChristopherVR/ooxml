import { readFileSync } from 'node:fs';
import { expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createEditSession } from '../edit/session';
import { pivotGrid } from './aggregate';
import * as columns from './columns';

const native = () =>
	readFileSync(
		new URL('../__fixtures__/excel-pivot-row-values-rows-columns.xlsx', import.meta.url),
	);

it('imports two real tabular row fields with Values last and preserves native parts', async () => {
	const bytes = native();
	const wb = await loadXlsx(bytes);
	expect(wb.warnings).toEqual([]);
	const pivot = wb.sheets[0]!.pivotTables![0]!;
	expect(pivot.secondRowField).toBe('Region');
	expect(pivot.valuesAxis).toBe('rows');
	const grid = pivotGrid(pivot);
	expect(grid.values).toHaveLength(23);
	expect(grid.rowLabelColumns).toBe(3);
	expect(grid.values.slice(0, 5)).toEqual([
		[null, null, null, 'Year', null, null],
		['Category', 'Region', 'Values', 2025, 2026, 'Grand Total'],
		['Fruit', 'East', 'Sum of Amount', 10, 7, 17],
		[null, null, 'Count of Quantity', 1, 1, 2],
		[null, null, 'Average of Amount', 10, 7, 8.5],
	]);
	expect(grid.values.at(-1)).toEqual(['Total Average of Amount', null, null, 7.5, 14, 11.4]);
	const before = await JSZip.loadAsync(bytes),
		after = await JSZip.loadAsync(await saveXlsx(wb));
	for (const part of Object.keys(before.files).filter(
		(part) => part.startsWith('xl/pivot') && !before.files[part]!.dir,
	))
		expect(await after.file(part)?.async('string')).toBe(await before.file(part)!.async('string'));
});

it('refreshes nested row measures from source records, saves and restores history', async () => {
	const wb = await loadXlsx(native());
	const edit = createEditSession(wb);
	const initial = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
	edit.setCellInput(1, 1, 2, '100');
	edit.refreshPivotTable(0, 'SalesPivot', { subtotals: false, repeatLabels: true });
	const grid = pivotGrid(wb.sheets[0]!.pivotTables![0]!);
	expect(grid.values).toHaveLength(17);
	expect(grid.values[3]).toEqual(['Fruit', 'East', 'Count of Quantity', 1, 1, 2]);
	expect(grid.values.at(-1)).toEqual(['Total Average of Amount', null, null, 52.5, 14, 29.4]);
	const loaded = await loadXlsx(await saveXlsx(wb));
	expect(loaded.warnings).toEqual([]);
	expect(pivotGrid(loaded.sheets[0]!.pivotTables![0]!).values).toEqual(grid.values);
	edit.undo();
	expect(pivotGrid(wb.sheets[0]!.pivotTables![0]!).values).toEqual(initial);
	edit.redo();
	expect(pivotGrid(wb.sheets[0]!.pivotTables![0]!).values).toEqual(grid.values);
});

it.each(['compact', 'outline'] as const)(
	'converts nested row measures to %s and restores history',
	async (layout) => {
		const wb = await loadXlsx(native());
		const edit = createEditSession(wb);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		edit.refreshPivotTable(0, pivot.name, { layout });
		const grid = pivotGrid(wb.sheets[0]!.pivotTables![0]!);
		expect(grid.values).toHaveLength(29);
		expect(grid.rowLabelColumns).toBe(layout === 'compact' ? 1 : 3);
		const loaded = await loadXlsx(await saveXlsx(wb));
		expect(loaded.warnings).toEqual([]);
		expect(pivotGrid(loaded.sheets[0]!.pivotTables![0]!).values).toEqual(grid.values);
		edit.undo();
		expect(wb.sheets[0]!.pivotTables![0]).toStrictEqual(pivot);
	},
);

it.each(['tabular', 'compact', 'outline'] as const)(
	'rejects oversized %s nested measure output before allocating headers or measure rows',
	async (layout) => {
		const wb = await loadXlsx(native());
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		delete pivot.columnField;
		pivot.layout = layout;
		pivot.additionalValues = Array.from({ length: 15 }, (_, i) => ({
			field: 'Amount',
			aggregate: 'sum' as const,
			caption: `Measure ${i}`,
		}));
		pivot.records = Array.from({ length: 35_000 }, (_, i) => ['Fruit', i, 10, 2, 2025]);
		const headers = vi.spyOn(columns, 'pivotColumnHeaders');
		try {
			expect(() => pivotGrid(pivot)).toThrow('too large');
			expect(headers).not.toHaveBeenCalled();
		} finally {
			headers.mockRestore();
		}
	},
);
