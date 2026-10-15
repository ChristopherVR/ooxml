import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/index';
import { pivotGrid } from './aggregate';

const fixture = (name: string) =>
	readFileSync(new URL(`../__fixtures__/excel-pivot-${name}.xlsx`, import.meta.url));
const axes = async (bytes: Uint8Array) => {
	const zip = await JSZip.loadAsync(bytes);
	const part = Object.keys(zip.files).find((n) => /pivotTables\/pivotTable\d+\.xml$/.test(n))!;
	const xml = await zip.file(part)!.async('string');
	return {
		items: (xml.match(/<items .*?<\/items>/g) ?? []).join(''),
		axes: xml.slice(xml.indexOf('<rowFields'), xml.indexOf('<dataFields')),
	};
};

const LAYOUTS = [
	['cols2-compact', ['Category', 'Year']],
	['cols2-tabular', ['Category', 'Year']],
	['cols2-outline', ['Category', 'Year']],
	['cols3-compact', ['Category', 'Region', 'Year']],
	['cols2-measures-compact', ['Category', 'Year']],
	['cols2-measures-tabular', ['Category', 'Year']],
	['cols2-rows2-compact-bottom', ['Category', 'Year']],
	['cols2-rows2-subtotals', ['Year', 'Product']],
] as const;

describe('several column fields', () => {
	it.each(LAYOUTS)('imports %s and reproduces its column items on refresh', async (name, cols) => {
		const bytes = fixture(name);
		const wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect([pivot.columnField, ...(pivot.moreColumnFields ?? [])]).toEqual(cols);
		const original = await JSZip.loadAsync(bytes);
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		for (const part of Object.keys(original.files).filter((n) => n.startsWith('xl/pivot')))
			expect(await saved.file(part)?.async('string'), part).toBe(
				await original.file(part)?.async('string'),
			);
		const before = await axes(bytes);
		createEditSession(wb).refreshPivotTable(0, pivot.name);
		const out = await saveXlsx(wb);
		const after = await axes(out);
		expect(after.items).toBe(before.items);
		const plain = (xml: string) => xml.replaceAll('<x v="0"/>', '<x/>');
		expect(plain(after.axes)).toBe(plain(before.axes));
		const reloaded = await loadXlsx(out);
		expect(reloaded.warnings).toEqual([]);
		expect(pivotGrid(reloaded.sheets[0]!.pivotTables![0]!).values).toEqual(pivotGrid(pivot).values);
	});

	it('lays out nested column headers with a subtotal column per group', async () => {
		const wb = await loadXlsx(fixture('cols2-compact'));
		const grid = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		expect(grid.slice(0, 3)).toEqual([
			['Data', 'Column Labels', null, null, null, null, null, null],
			[null, 'Fruit', null, 'Fruit Total', 'Veg', null, 'Veg Total', 'Grand Total'],
			['Row Labels', 2025, 2026, null, 2025, 2026, null, null],
		]);
		expect(grid[3]).toEqual(['East', 10, 10.5, 20.5, 5, 9, 14, 34.5]);
		expect(grid.at(-1)).toEqual(['Grand Total', 30, 12.75, 42.75, 20, 13, 33, 75.75]);
	});

	it('names the fields in the tabular form and repeats measures under each group', async () => {
		const tabular = pivotGrid(
			(await loadXlsx(fixture('cols2-tabular'))).sheets[0]!.pivotTables![0]!,
		);
		expect(tabular.values[0]).toEqual(['Data', 'Category', 'Year', null, null, null, null, null]);
		const several = pivotGrid(
			(await loadXlsx(fixture('cols2-measures-compact'))).sheets[0]!.pivotTables![0]!,
		);
		expect(several.values[1]).toEqual([
			null,
			'Fruit',
			null,
			null,
			null,
			'Fruit Data',
			'Fruit Count of Amount',
			'Veg',
			null,
			null,
			null,
			'Veg Data',
			'Veg Count of Amount',
			'Total Data',
			'Total Count of Amount',
		]);
		expect(several.values[3]![1]).toBe('Data');
	});

	it('is chosen when creating a report, and changes through a refresh with undo', async () => {
		const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
		const edit = createEditSession(wb);
		[
			['A', 'B', 'C', 'N'],
			['a1', 'b1', 'c1', 1],
			['a1', 'b2', 'c1', 2],
			['a2', 'b1', 'c2', 4],
			['a2', 'b2', 'c2', 8],
		].forEach((row, r) => row.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
		const name = edit.createPivotTable(
			1,
			{ row: 0, col: 0 },
			{
				source: 'Data!A1:D5',
				rowField: 'C',
				columnField: 'A',
				moreColumnFields: ['B'],
				valueField: 'N',
				layout: 'compact',
			},
		);
		const pivot = () => wb.sheets[1]!.pivotTables![0]!;
		expect(pivotGrid(pivot()).values.map((line) => line[0])).toEqual([
			'Sum of N',
			null,
			'Row Labels',
			'c1',
			'c2',
			'Grand Total',
		]);
		const reloaded = await loadXlsx(await saveXlsx(wb));
		expect(reloaded.warnings).toEqual([]);
		expect(reloaded.sheets[1]!.pivotTables![0]).toMatchObject({
			columnField: 'A',
			moreColumnFields: ['B'],
		});
		edit.refreshPivotTable(1, name, { moreColumnFields: [] });
		expect(pivot().moreColumnFields).toBeUndefined();
		expect(pivotGrid(pivot()).values).toHaveLength(5);
		edit.undo();
		expect(pivot().moreColumnFields).toEqual(['B']);
		// Dropping the first column field drops the ones after it.
		edit.refreshPivotTable(1, name, { columnField: '' });
		expect(pivot().columnField).toBeUndefined();
		expect(pivot().moreColumnFields).toBeUndefined();
	});

	it('rejects column fields that cannot form a report', () => {
		const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
		const edit = createEditSession(wb);
		[
			['A', 'B', 'N'],
			['a1', 'b1', 1],
			['a2', 'b2', 2],
		].forEach((row, r) => row.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
		const base = { source: 'Data!A1:C3', rowField: 'A', valueField: 'N' };
		const create = (extra: object) => () =>
			edit.createPivotTable(1, { row: 0, col: 0 }, { ...base, ...extra });
		expect(create({ columnField: 'B', moreColumnFields: ['B'] })).toThrow(
			'row and column fields must differ',
		);
		expect(create({ columnField: 'B', moreColumnFields: ['A'] })).toThrow(
			'row and column fields must differ',
		);
		expect(
			create({
				columnField: 'B',
				moreColumnFields: ['N'],
				valuesAxis: 'rows',
				additionalValues: [{ field: 'N' }],
			}),
		).toThrow('Values on rows support one column field');
		expect(
			create({ pageFilter: { field: 'N' }, columnField: 'B', moreColumnFields: ['N'] }),
		).toThrow();
	});

	it('groups a date field on the columns by several units', async () => {
		const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
		const edit = createEditSession(wb);
		[
			['Date', 'Region', 'N'],
			[45292, 'East', 1],
			[45323, 'West', 2],
			[45658, 'East', 4],
		].forEach((row, r) => row.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
		edit.applyStyle(0, [{ start: { row: 1, col: 0 }, end: { row: 3, col: 0 } }], {
			numFmt: 'yyyy-mm-dd',
		});
		edit.createPivotTable(
			1,
			{ row: 0, col: 0 },
			{
				source: 'Data!A1:C4',
				rowField: 'Region',
				columnField: 'Date',
				valueField: 'N',
				dateGroups: [{ field: 'Date', by: ['months', 'years'] }],
			},
		);
		const pivot = wb.sheets[1]!.pivotTables![0]!;
		expect(pivot).toMatchObject({
			columnField: 'Years (Date)',
			moreColumnFields: ['Months (Date)'],
		});
		const grid = pivotGrid(pivot).values;
		expect(grid[1]).toEqual([
			null,
			'2024',
			null,
			'2024 Total',
			'2025',
			'2025 Total',
			'Grand Total',
		]);
		expect(grid.at(-1)).toEqual(['Grand Total', 1, 2, 3, 4, 4, 7]);
		const reloaded = await loadXlsx(await saveXlsx(wb));
		expect(reloaded.warnings).toEqual([]);
		expect(pivotGrid(reloaded.sheets[1]!.pivotTables![0]!).values).toEqual(grid);
	});
});
