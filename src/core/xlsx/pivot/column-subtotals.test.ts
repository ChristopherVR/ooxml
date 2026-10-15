import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/index';
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
		// Only the subtotal flags: Excel's own attribute order and defaults differ from ours.
		flags: (xml.match(/<pivotField [^>]*>/g) ?? [])
			.map((tag) => (tag.match(/\b(?!default)\w+Subtotal="1"/g) ?? []).join(' '))
			.join('|'),
		axes: xml.slice(xml.indexOf('<rowFields'), xml.indexOf('<dataFields')),
	};
};

const CASES = [
	['colsubs-none-compact', [{ field: 'Category', functions: [] }]],
	['colsubs-sum-average-compact', [{ field: 'Category', functions: ['sum', 'average'] }]],
	['colsubs-sum-average-tabular', [{ field: 'Category', functions: ['sum', 'average'] }]],
	['colsubs-count-outline', [{ field: 'Category', functions: ['count'] }]],
	[
		'colsubs-three-levels',
		[
			{ field: 'Category', functions: ['sum', 'average'] },
			{ field: 'Region', functions: [] },
		],
	],
	['colsubs-measures-compact', [{ field: 'Category', functions: ['sum', 'average'] }]],
	['colsubs-measures-other', [{ field: 'Category', functions: ['sum', 'count'] }]],
	[
		'colsubs-rows-subtotals',
		[
			{ field: 'Category', functions: ['sum', 'average'] },
			{ field: 'Year', functions: [] },
		],
	],
	[
		'colsubs-rows-custom',
		[
			{ field: 'Category', functions: ['sum', 'average'] },
			{ field: 'Year', functions: ['sum', 'count'] },
		],
	],
	[
		'colsubs-rows-custom-compact',
		[
			{ field: 'Category', functions: ['sum', 'average'] },
			{ field: 'Year', functions: ['sum', 'count'] },
		],
	],
	['colsubs-rows-total', [{ field: 'Year', functions: ['sum', 'count'] }]],
	['colsubs-rows-auto', [{ field: 'Category', functions: ['sum', 'average'] }]],
	[
		'colsubs-rows-same',
		[
			{ field: 'Category', functions: ['sum', 'count'] },
			{ field: 'Year', functions: ['sum', 'count'] },
		],
	],
] as const;

describe('column field subtotals', () => {
	it.each(CASES)('imports %s and reproduces its parts', async (name, subtotals) => {
		const bytes = fixture(name);
		const wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.fieldSubtotals).toEqual(subtotals);
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
		expect(after.flags).toBe(before.flags);
		const plain = (xml: string) => xml.replaceAll('<x v="0"/>', '<x/>');
		expect(plain(after.axes)).toBe(plain(before.axes));
		const reloaded = await loadXlsx(out);
		expect(reloaded.warnings).toEqual([]);
		expect(pivotGrid(reloaded.sheets[0]!.pivotTables![0]!).values).toEqual(pivotGrid(pivot).values);
	});

	it('adds one subtotal column per function after a group', async () => {
		const wb = await loadXlsx(fixture('colsubs-sum-average-compact'));
		const grid = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		expect(grid.slice(0, 3)).toEqual([
			['Data', 'Column Labels', null, null, null, null, null, null, null, null],
			[
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
			],
			['Row Labels', 2025, 2026, null, null, 2025, 2026, null, null, null],
		]);
		expect(grid[3]).toEqual(['East', 10, 10.5, 20.5, 6.833333333333333, 5, 9, 14, 7, 34.5]);
	});

	it('removes the subtotal columns for None', async () => {
		const wb = await loadXlsx(fixture('colsubs-none-compact'));
		const grid = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		expect(grid[1]).toEqual([null, 'Fruit', null, 'Veg', null, 'Grand Total']);
	});

	it('names the source field of a custom subtotal when there are several measures', async () => {
		const wb = await loadXlsx(fixture('colsubs-measures-other'));
		const grid = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		expect(grid[1]!.slice(5, 9)).toEqual([
			'Fruit Sum of Amount',
			'Fruit Sum of Mixed',
			'Fruit Count of Amount',
			'Fruit Count of Mixed',
		]);
		expect(grid[4]!.slice(5, 9)).toEqual([20.5, 3, 3, 3]);
	});

	it('leaves a cell empty where a row and a column subtotal use different functions', async () => {
		const wb = await loadXlsx(fixture('colsubs-rows-custom'));
		const grid = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		expect(grid[5]).toEqual([
			'Fruit Sum',
			null,
			30,
			null,
			30,
			null,
			3,
			null,
			null,
			7.5,
			2.25,
			12.75,
			null,
			42.75,
		]);
		expect(grid[6]).toEqual([
			'Fruit Average',
			null,
			15,
			null,
			null,
			null,
			3,
			null,
			null,
			7.5,
			2.25,
			null,
			null,
			8.55,
		]);
	});

	it('changes a column field subtotal through the edit session and writes it back', async () => {
		const wb = await loadXlsx(fixture('colsubs-none-compact'));
		const session = createEditSession(wb);
		session.refreshPivotTable(0, 'ColumnsPivot', {
			fieldSubtotals: [{ field: 'Category', functions: ['max', 'sum'] }],
		});
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.fieldSubtotals).toEqual([{ field: 'Category', functions: ['sum', 'max'] }]);
		expect(pivotGrid(pivot).values[1]).toEqual([
			null,
			'Fruit',
			null,
			'Fruit Sum',
			'Fruit Max',
			'Veg',
			null,
			'Veg Sum',
			'Veg Max',
			'Grand Total',
		]);
		const reloaded = await loadXlsx(await saveXlsx(wb));
		expect(reloaded.warnings).toEqual([]);
		expect(reloaded.sheets[0]!.pivotTables![0]!.fieldSubtotals).toEqual(pivot.fieldSubtotals);
	});

	it('drops the subtotal setting of a field that leaves the axes', async () => {
		const wb = await loadXlsx(fixture('colsubs-sum-average-compact'));
		createEditSession(wb).refreshPivotTable(0, 'ColumnsPivot', {
			columnField: 'Year',
			moreColumnFields: [],
		});
		expect(wb.sheets[0]!.pivotTables![0]!.fieldSubtotals).toBeUndefined();
	});

	it('rejects a subtotal for a field that is not on an axis', async () => {
		const wb = await loadXlsx(fixture('colsubs-none-compact'));
		expect(() =>
			createEditSession(wb).refreshPivotTable(0, 'ColumnsPivot', {
				fieldSubtotals: [{ field: 'Product', functions: ['sum'] }],
			}),
		).toThrow('Choose valid field subtotals.');
	});
});
