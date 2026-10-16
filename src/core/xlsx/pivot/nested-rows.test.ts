import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { getCell } from '../cells';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/index';
import { pivotGrid } from './aggregate';

const fixture = (name: string) =>
	readFileSync(new URL(`../__fixtures__/excel-pivot-${name}.xlsx`, import.meta.url));
const rowItems = async (bytes: Uint8Array, part?: string) => {
	const zip = await JSZip.loadAsync(bytes);
	const name =
		part ?? Object.keys(zip.files).find((n) => /pivotTables\/pivotTable\d+\.xml$/.test(n))!;
	const xml = await zip.file(name)!.async('string');
	return {
		xml,
		fields: (xml.match(/<items .*?<\/items>/g) ?? []).join(''),
		rows: xml.slice(xml.indexOf('<rowFields'), xml.indexOf('<colItems')),
	};
};
const labels = (wb: Awaited<ReturnType<typeof loadXlsx>>, column = 0) =>
	pivotGrid(wb.sheets[0]!.pivotTables![0]!).values.map((line) => line[column]);

const LAYOUTS = [
	'rows3-compact-top',
	'rows3-compact-bottom',
	'rows3-outline-top',
	'rows3-outline-bottom',
	'rows3-outline-top-repeat',
	'rows3-outline-bottom-repeat',
	'rows3-tabular',
	'rows3-tabular-repeat',
	'rows3-tabular-columns',
] as const;

describe('three row fields', () => {
	it.each(LAYOUTS)(
		'imports the native %s layout and reproduces its items on refresh',
		async (name) => {
			const bytes = fixture(name);
			const wb = await loadXlsx(bytes);
			expect(wb.warnings).toEqual([]);
			const pivot = wb.sheets[0]!.pivotTables![0]!;
			expect(pivot).toMatchObject({
				rowField: 'Category',
				secondRowField: 'Region',
				moreRowFields: ['Product'],
			});
			// Untouched definitions are saved verbatim.
			const zip = await JSZip.loadAsync(bytes);
			const saved = await JSZip.loadAsync(await saveXlsx(wb));
			for (const part of Object.keys(zip.files).filter(
				(n) => n.startsWith('xl/pivot') && !zip.files[n]!.dir,
			))
				expect(await saved.file(part)?.async('string'), part).toBe(
					await zip.file(part)!.async('string'),
				);
			// A refresh rewrites them from the model: same items and row items as Excel wrote.
			const original = await rowItems(bytes);
			createEditSession(wb).refreshPivotTable(0, pivot.name);
			const rewritten = await rowItems(await saveXlsx(wb));
			expect(rewritten.fields).toBe(original.fields);
			expect(rewritten.rows.replaceAll('<x v="0"/>', '<x/>')).toBe(
				original.rows.replaceAll('<x v="0"/>', '<x/>'),
			);
			const reloaded = await loadXlsx(await saveXlsx(wb));
			expect(reloaded.warnings).toEqual([]);
			expect(pivotGrid(reloaded.sheets[0]!.pivotTables![0]!).values).toEqual(
				pivotGrid(pivot).values,
			);
		},
	);

	it('nests subtotals at every level in the compact form', async () => {
		const wb = await loadXlsx(fixture('rows3-compact-bottom'));
		const sheet = wb.sheets[0]!;
		expect(labels(wb)).toEqual([
			'Row Labels',
			'Fruit',
			'East',
			'Apple',
			'Pear',
			'East Total',
			'West',
			'Apple',
			'Plum',
			'West Total',
			'Fruit Total',
			'Veg',
			'East',
			'Carrot',
			'Leek',
			'East Total',
			'West',
			'Carrot',
			'West Total',
			'Veg Total',
			'Grand Total',
		]);
		expect(getCell(sheet, 12, 1)?.value).toBe(42.75);
		const pivot = sheet.pivotTables![0]!;
		expect(pivotGrid(pivot).indents?.slice(0, 7)).toEqual([0, 0, 1, 2, 2, 1, 1]);
	});

	it('prints an ancestor label on the first row of its group in the tabular form', async () => {
		const wb = await loadXlsx(fixture('rows3-tabular'));
		const grid = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		expect(grid.slice(1, 8)).toEqual([
			['Fruit', 'East', 'Apple', 13],
			[null, null, 'Pear', 7.5],
			[null, 'East Total', null, 20.5],
			[null, 'West', 'Apple', 20],
			[null, null, 'Plum', 2.25],
			[null, 'West Total', null, 22.25],
			['Fruit Total', null, null, 42.75],
		]);
	});

	it('repeats every ancestor label, subtotal rows included', async () => {
		const wb = await loadXlsx(fixture('rows3-tabular-repeat'));
		const grid = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		expect(grid.slice(3, 6)).toEqual([
			['Fruit', 'East Total', null, 20.5],
			['Fruit', 'West', 'Apple', 20],
			['Fruit', 'West', 'Plum', 2.25],
		]);
	});

	it('splits every level by the column field', async () => {
		const wb = await loadXlsx(fixture('rows3-tabular-columns'));
		const grid = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		expect(grid[1]).toEqual(['Category', 'Region', 'Product', 2025, 2026, 'Grand Total']);
		expect(grid.at(-2)).toEqual(['Veg Total', null, null, 20, 13, 33]);
		expect(grid.at(-1)).toEqual(['Grand Total', null, null, 50, 25.75, 75.75]);
	});
});

describe('creating and editing many row fields', () => {
	const setup = () => {
		const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
		const edit = createEditSession(wb);
		[
			['A', 'B', 'C', 'D', 'N'],
			['a1', 'b1', 'c1', 'd1', 1],
			['a1', 'b1', 'c1', 'd2', 2],
			['a1', 'b1', 'c2', 'd1', 4],
			['a1', 'b2', 'c1', 'd1', 8],
			['a2', 'b1', 'c1', 'd1', 16],
		].forEach((row, r) => row.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
		return { wb, edit };
	};
	const options = { source: 'Data!A1:E6', rowField: 'A', secondRowField: 'B', valueField: 'N' };

	it('computes four nested levels with a subtotal under each', () => {
		const { wb, edit } = setup();
		edit.createPivotTable(1, { row: 0, col: 0 }, { ...options, moreRowFields: ['C', 'D'] });
		const grid = pivotGrid(wb.sheets[1]!.pivotTables![0]!).values;
		expect(grid[0]).toEqual(['A', 'B', 'C', 'D', 'Sum of N']);
		expect(grid.filter((line) => String(line[1]).endsWith('Total'))).toEqual([
			[null, 'b1 Total', null, null, 7],
			[null, 'b2 Total', null, null, 8],
			[null, 'b1 Total', null, null, 16],
		]);
		expect(grid.at(-2)).toEqual(['a2 Total', null, null, null, 16]);
		expect(grid.at(-1)).toEqual(['Grand Total', null, null, null, 31]);
		// Subtotals sit under the groups of the outer three fields.
		expect(grid.filter((line) => String(line[2]).endsWith('Total'))).toHaveLength(4);
	});

	it('adds and removes row fields through a refresh, with undo', async () => {
		const { wb, edit } = setup();
		const name = edit.createPivotTable(1, { row: 0, col: 0 }, options);
		const grid = () => pivotGrid(wb.sheets[1]!.pivotTables![0]!).values;
		const before = grid();
		edit.refreshPivotTable(1, name, { moreRowFields: ['C', 'D'] });
		expect(wb.sheets[1]!.pivotTables![0]!.moreRowFields).toEqual(['C', 'D']);
		expect(grid()[0]).toEqual(['A', 'B', 'C', 'D', 'Sum of N']);
		const saved = await loadXlsx(await saveXlsx(wb));
		expect(saved.warnings).toEqual([]);
		expect(pivotGrid(saved.sheets[1]!.pivotTables![0]!).values).toEqual(grid());
		edit.refreshPivotTable(1, name, { moreRowFields: [] });
		expect(grid()).toEqual(before);
		edit.undo();
		expect(grid()[0]).toEqual(['A', 'B', 'C', 'D', 'Sum of N']);
		// Dropping the second row field drops the ones after it.
		edit.refreshPivotTable(1, name, { secondRowField: '' });
		expect(wb.sheets[1]!.pivotTables![0]!.moreRowFields).toBeUndefined();
	});

	it('rejects row fields that cannot form a report', () => {
		const { edit } = setup();
		const create = (extra: object) => () =>
			edit.createPivotTable(1, { row: 0, col: 0 }, { ...options, ...extra });
		expect(create({ secondRowField: undefined, moreRowFields: ['C'] })).not.toThrow();
		expect(create({ moreRowFields: ['C', 'C'] })).toThrow('row and column fields must differ');
		expect(create({ moreRowFields: ['A'] })).toThrow('row and column fields must differ');
		expect(create({ moreRowFields: ['Nope'] })).toThrow('Choose valid PivotTable fields');
		// Values on rows take any number of row fields.
		expect(() =>
			edit.createPivotTable(
				1,
				{ row: 60, col: 0 },
				{
					...options,
					moreRowFields: ['C'],
					valuesAxis: 'rows',
					additionalValues: [{ field: 'N', aggregate: 'count' }],
				},
			),
		).not.toThrow();
	});
});
