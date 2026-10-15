import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/index';
import { pivotGrid } from './aggregate';
import { aggregateValues } from './measures';
import { fieldSubtotalFunctions, PIVOT_SUBTOTAL_ORDER } from './subtotals';

const fixture = (name: string) =>
	readFileSync(new URL(`../__fixtures__/excel-pivot-subtotals-${name}.xlsx`, import.meta.url));
const tableXml = async (bytes: Uint8Array) => {
	const zip = await JSZip.loadAsync(bytes);
	const part = Object.keys(zip.files).find((n) => /pivotTables\/pivotTable\d+\.xml$/.test(n))!;
	return zip.file(part)!.async('string');
};

describe('summary functions', () => {
	const records = [[10], [7.5], [20], [2.25], [3], [null], ['x']];
	it('match Excel on the same values', () => {
		const sample = (type: Parameters<typeof aggregateValues>[2]) =>
			aggregateValues(records, 0, type);
		expect(sample('sum')).toBe(42.75);
		expect(sample('count')).toBe(6);
		expect(sample('countNums')).toBe(5);
		expect(sample('product')).toBe(10125);
		expect(sample('var')).toBeCloseTo(51.2, 12);
		expect(sample('varp')).toBeCloseTo(40.96, 12);
		expect(sample('stdDev')).toBeCloseTo(7.1554175279993268, 12);
		expect(sample('stdDevp')).toBeCloseTo(6.4, 12);
	});
	it('divide by zero when there are too few numbers', () => {
		expect(aggregateValues([[5]], 0, 'var')).toEqual({ error: '#DIV/0!' });
		expect(aggregateValues([[5]], 0, 'stdDev')).toEqual({ error: '#DIV/0!' });
		expect(aggregateValues([[5]], 0, 'varp')).toBe(0);
		expect(aggregateValues([['x']], 0, 'varp')).toEqual({ error: '#DIV/0!' });
		expect(aggregateValues([['x']], 0, 'countNums')).toBe(0);
	});
});

describe('custom subtotals', () => {
	it('imports every function Excel offers and reproduces its rows', async () => {
		const bytes = fixture('all');
		const wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.fieldSubtotals).toEqual([
			{ field: 'Category', functions: [...PIVOT_SUBTOTAL_ORDER] },
		]);
		const grid = pivotGrid(pivot).values;
		const fruit = grid.filter((line) => String(line[0]).startsWith('Fruit '));
		expect(fruit.map((line) => [line[0], line[2]])).toEqual([
			['Fruit Sum', 42.75],
			['Fruit Count', 5],
			['Fruit Average', 8.55],
			['Fruit Max', 20],
			['Fruit Min', 2.25],
			['Fruit Product', 10125],
			['Fruit Count', 5],
			['Fruit StdDev', 7.1554175279993268],
			['Fruit StdDevp', 6.4],
			['Fruit Var', 51.2],
			['Fruit Varp', 40.96],
		]);
		expect(grid.at(-1)).toEqual(['Grand Total', null, 75.75]);
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		const original = await JSZip.loadAsync(bytes);
		for (const part of Object.keys(original.files).filter((n) => n.startsWith('xl/pivot')))
			expect(await saved.file(part)?.async('string')).toBe(
				await original.file(part)?.async('string'),
			);
	});

	it('rewrites the field flags and subtotal items the way Excel does', async () => {
		const wb = await loadXlsx(fixture('some-compact'));
		createEditSession(wb).refreshPivotTable(0, 'FieldsPivot');
		const xml = await tableXml(await saveXlsx(wb));
		expect(xml).toContain('avgSubtotal="1"');
		expect(xml).toContain('minSubtotal="1"');
		expect(xml).toContain(
			'<items count="4"><item x="0"/><item x="1"/><item t="avg"/><item t="min"/></items>',
		);
		expect(xml).toContain(
			'<items count="4"><item x="0"/><item x="1"/><item t="sum"/><item t="avg"/></items>',
		);
		// Custom subtotals replace the automatic one on those fields only.
		expect(xml).toContain('<item t="default"/>');
		expect(xml.match(/<i t="avg"/g)).toHaveLength(6);
	});

	it('computes from the source values whatever the value field summarises', async () => {
		const wb = await loadXlsx(fixture('count-field'));
		const grid = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		expect(grid.slice(1, 5)).toEqual([
			['Fruit', 'East', 3],
			[null, 'West', 2],
			['Fruit Sum', null, 42.75],
			['Fruit Average', null, 8.55],
		]);
	});

	it('shows no subtotal for a field that has none', async () => {
		const wb = await loadXlsx(fixture('none-compact'));
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.fieldSubtotals).toEqual([{ field: 'Region', functions: [] }]);
		const grid = pivotGrid(pivot).values;
		expect(grid.map((line) => line[0])).not.toContain('East Total');
		expect(grid[1]).toEqual(['Fruit', 42.75]);
		expect(grid[2]).toEqual(['East', null]);
	});

	it('is chosen per field when creating or refreshing a report', async () => {
		const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
		const edit = createEditSession(wb);
		[
			['G', 'H', 'N'],
			['a', 'x', 1],
			['a', 'y', 5],
			['b', 'x', 9],
		].forEach((row, r) => row.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
		const name = edit.createPivotTable(
			1,
			{ row: 0, col: 0 },
			{
				source: 'Data!A1:C4',
				rowField: 'G',
				secondRowField: 'H',
				valueField: 'N',
				fieldSubtotals: [{ field: 'G', functions: ['max', 'sum', 'sum'].slice(0, 2) as never }],
			},
		);
		const pivot = () => wb.sheets[1]!.pivotTables![0]!;
		// Functions are listed in Excel's order, not the order they were chosen in.
		expect(pivot().fieldSubtotals).toEqual([{ field: 'G', functions: ['sum', 'max'] }]);
		expect(fieldSubtotalFunctions(pivot(), 'G')).toEqual(['sum', 'max']);
		expect(fieldSubtotalFunctions(pivot(), 'H')).toBe('auto');
		expect(pivotGrid(pivot()).values.map((line) => line[0])).toEqual([
			'G',
			'a',
			null,
			'a Sum',
			'a Max',
			'b',
			'b Sum',
			'b Max',
			'Grand Total',
		]);
		const reloaded = await loadXlsx(await saveXlsx(wb));
		expect(reloaded.warnings).toEqual([]);
		expect(reloaded.sheets[1]!.pivotTables![0]!.fieldSubtotals).toEqual([
			{ field: 'G', functions: ['sum', 'max'] },
		]);
		edit.refreshPivotTable(1, name, { fieldSubtotals: [] });
		expect(pivot().fieldSubtotals).toBeUndefined();
		expect(pivotGrid(pivot()).values.map((line) => line[0])).toContain('a Total');
		edit.refreshPivotTable(1, name, { fieldSubtotals: [{ field: 'G', functions: [] }] });
		expect(pivotGrid(pivot()).values.map((line) => line[0])).not.toContain('a Total');
		edit.undo();
		expect(pivotGrid(pivot()).values.map((line) => line[0])).toContain('a Total');
	});

	it('rejects unknown fields and functions, and Values on rows', () => {
		const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
		const edit = createEditSession(wb);
		[
			['G', 'H', 'N'],
			['a', 'x', 1],
			['b', 'y', 2],
		].forEach((row, r) => row.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
		const base = { source: 'Data!A1:C3', rowField: 'G', secondRowField: 'H', valueField: 'N' };
		const create = (extra: object) => () =>
			edit.createPivotTable(1, { row: 0, col: 0 }, { ...base, ...extra });
		expect(create({ fieldSubtotals: [{ field: 'N', functions: ['sum'] }] })).toThrow(
			'Choose valid field subtotals',
		);
		expect(create({ fieldSubtotals: [{ field: 'G', functions: ['median'] }] })).toThrow(
			'Choose valid field subtotals',
		);
		expect(
			create({
				fieldSubtotals: [{ field: 'G', functions: ['sum'] }],
				valuesAxis: 'rows',
				additionalValues: [{ field: 'N' }],
			}),
		).toThrow('Field subtotals require Values on columns');
	});
});
