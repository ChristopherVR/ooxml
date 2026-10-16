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
const parts = async (bytes: Uint8Array) => {
	const zip = await JSZip.loadAsync(bytes);
	const part = Object.keys(zip.files).find((n) => /pivotTables\/pivotTable\d+\.xml$/.test(n))!;
	const xml = await zip.file(part)!.async('string');
	const plain = (text: string) => text.replaceAll('<x v="0"/>', '<x/>');
	return {
		items: (xml.match(/<items .*?<\/items>/g) ?? []).join(''),
		flags: (xml.match(/<pivotField [^>]*>/g) ?? [])
			.map((tag) => (tag.match(/\b(?!default)\w+Subtotal="1"/g) ?? []).join(' '))
			.join('|'),
		location: xml.match(/<location [^>]*>/)![0].replace(/ ref="[^"]*"/, ''),
		position: xml.match(/dataPosition="(\d+)"/)?.[1],
		axes: plain(xml.slice(xml.indexOf('<rowFields'), xml.indexOf('<dataFields'))),
	};
};

/** Name, row fields with Values at its place, column fields. */
const CASES = [
	['rv3-last-compact-top', ['Category', 'Region', 'Product', 'Values'], []],
	['rv3-last-compact-bottom', ['Category', 'Region', 'Product', 'Values'], []],
	['rv3-last-outline-top', ['Category', 'Region', 'Product', 'Values'], []],
	['rv3-last-outline-bottom', ['Category', 'Region', 'Product', 'Values'], []],
	['rv3-last-outline-top-repeat', ['Category', 'Region', 'Product', 'Values'], []],
	['rv3-last-tabular', ['Category', 'Region', 'Product', 'Values'], []],
	['rv3-first-compact-top', ['Values', 'Category', 'Region', 'Product'], []],
	['rv3-first-compact-bottom', ['Values', 'Category', 'Region', 'Product'], []],
	['rv3-first-outline-top', ['Values', 'Category', 'Region', 'Product'], []],
	['rv3-first-outline-bottom', ['Values', 'Category', 'Region', 'Product'], []],
	['rv3-first-tabular', ['Values', 'Category', 'Region', 'Product'], []],
	['rv3-first-tabular-repeat', ['Values', 'Category', 'Region', 'Product'], []],
	['rv3-second-compact-top', ['Category', 'Values', 'Region', 'Product'], []],
	['rv3-second-outline-bottom', ['Category', 'Values', 'Region', 'Product'], []],
	['rv3-second-outline-top-repeat', ['Category', 'Values', 'Region', 'Product'], []],
	['rv3-second-tabular', ['Category', 'Values', 'Region', 'Product'], []],
	['rv3-third-compact-top', ['Category', 'Region', 'Values', 'Product'], []],
	['rv3-third-outline-top', ['Category', 'Region', 'Values', 'Product'], []],
	['rv3-third-tabular', ['Category', 'Region', 'Values', 'Product'], []],
	['rv3-third-tabular-repeat', ['Category', 'Region', 'Values', 'Product'], []],
	['rv2-cols2-last-compact', ['Category', 'Region', 'Values'], ['Year', 'Product']],
	['rv2-cols2-first-tabular', ['Values', 'Category', 'Region'], ['Year', 'Product']],
	['rv2-cols2-second-outline', ['Category', 'Values', 'Region'], ['Year', 'Product']],
	['rv1-cols2-compact', ['Region', 'Values'], ['Category', 'Year']],
	['rv1-cols2-tabular', ['Values', 'Region'], ['Category', 'Year']],
	['rv1-first-compact', ['Values', 'Region'], []],
	['rv1-first-outline-cols', ['Values', 'Region'], ['Category']],
	['rvs-last-compact', ['Category', 'Region', 'Product', 'Values'], []],
	['rvs-first-compact-top', ['Values', 'Category', 'Region', 'Product'], []],
	['rvs-first-compact-bottom', ['Values', 'Category', 'Region', 'Product'], []],
	['rvs-second-tabular', ['Category', 'Values', 'Region', 'Product'], []],
	['rvs-third-outline', ['Category', 'Region', 'Values', 'Product'], []],
	['rvs-cols-compact', ['Region', 'Values'], ['Category', 'Year']],
	['rvs-rows-cols-tabular', ['Category', 'Region', 'Values'], ['Year', 'Product']],
	['rvs-first-tabular-cols', ['Values', 'Category', 'Region'], ['Year', 'Product']],
	['rvs-none-last', ['Category', 'Region', 'Values'], []],
	['rvs-none-first', ['Values', 'Category', 'Region'], []],
] as const;

describe('Values on rows with any number of row and column fields', () => {
	it.each(CASES)('imports %s and reproduces its parts', async (name, rows, columns) => {
		const bytes = fixture(name);
		const wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		const fields = rows.filter((field) => field !== 'Values');
		expect([pivot.rowField, pivot.secondRowField, ...(pivot.moreRowFields ?? [])]).toEqual(
			fields.length === 1 ? [fields[0], undefined] : fields,
		);
		expect([pivot.columnField, ...(pivot.moreColumnFields ?? [])].filter(Boolean)).toEqual(columns);
		expect(pivot.valuesAxis).toBe('rows');
		const original = await JSZip.loadAsync(bytes);
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		for (const part of Object.keys(original.files).filter((n) => n.startsWith('xl/pivot')))
			expect(await saved.file(part)?.async('string'), part).toBe(
				await original.file(part)?.async('string'),
			);
		const before = await parts(bytes);
		createEditSession(wb).refreshPivotTable(0, pivot.name);
		const out = await saveXlsx(wb);
		expect(await parts(out)).toEqual(before);
		const reloaded = await loadXlsx(out);
		expect(reloaded.warnings).toEqual([]);
		expect(pivotGrid(reloaded.sheets[0]!.pivotTables![0]!).values).toEqual(pivotGrid(pivot).values);
	});

	it('numbers the Values position as the fields before it', async () => {
		const read = async (name: string) =>
			(await loadXlsx(fixture(name))).sheets[0]!.pivotTables![0]!.valuesPosition;
		expect(await read('rv3-last-compact-top')).toBeUndefined();
		expect(await read('rv3-first-compact-top')).toBe('first');
		expect(await read('rv3-second-compact-top')).toBe(1);
		expect(await read('rv3-third-compact-top')).toBe(2);
		expect(await read('rv2-cols2-second-outline')).toBe('between');
	});

	it('places subtotal rows by where Values sit', async () => {
		const labels = async (name: string) =>
			pivotGrid((await loadXlsx(fixture(name))).sheets[0]!.pivotTables![0]!).values.map(
				(line) => line[0],
			);
		// Values below the field: one row per measure after the group.
		expect(await labels('rv3-last-compact-top')).toContain('East Data');
		expect(await labels('rv3-last-compact-top')).toContain('Fruit Count of Amount');
		// Values above: a Total row, or the group's own row when the subtotal sits above.
		expect(await labels('rv3-first-compact-bottom')).toContain('East Total');
		expect(await labels('rv3-first-compact-top')).not.toContain('East Total');
	});

	it('expands a custom subtotal into one row per measure below the field', async () => {
		const wb = await loadXlsx(fixture('rvs-last-compact'));
		const rows = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		expect(rows.filter((line) => String(line[0]).startsWith('Fruit '))).toEqual([
			['Fruit Sum of Amount', 42.75],
			['Fruit Sum of Amount', 42.75],
			['Fruit Average of Amount', 8.55],
			['Fruit Average of Amount', 8.55],
		]);
		const above = await loadXlsx(fixture('rvs-first-compact-top'));
		expect(
			pivotGrid(above.sheets[0]!.pivotTables![0]!)
				.values.filter((line) => String(line[0]).startsWith('Fruit '))
				.map((line) => line[0]),
		).toEqual(['Fruit Sum', 'Fruit Average', 'Fruit Sum', 'Fruit Average']);
	});

	it('crosses row and column subtotals only where their functions agree', async () => {
		const wb = await loadXlsx(fixture('rvs-rows-cols-tabular'));
		const grid = pivotGrid(wb.sheets[0]!.pivotTables![0]!).values;
		const sum = grid.find((line) => line[0] === 'Fruit Sum of Amount')!;
		const average = grid.find((line) => line[0] === 'Fruit Average of Amount')!;
		expect(sum.slice(3, 7)).toEqual([30, null, 30, null]);
		expect(average.slice(3, 7)).toEqual([15, null, null, null]);
	});

	it('creates and edits reports with Values anywhere among several fields', () => {
		const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
		const edit = createEditSession(wb);
		[
			['A', 'B', 'C', 'D', 'N'],
			['a1', 'b1', 'c1', 'd1', 1],
			['a1', 'b1', 'c2', 'd2', 2],
			['a2', 'b2', 'c1', 'd1', 4],
		].forEach((row, r) => row.forEach((v, c) => edit.setCellInput(0, r, c, String(v))));
		edit.createPivotTable(
			1,
			{ row: 0, col: 0 },
			{
				source: 'Data!A1:E4',
				rowField: 'A',
				secondRowField: 'B',
				moreRowFields: ['C'],
				columnField: 'D',
				valueField: 'N',
				additionalValues: [{ field: 'N', aggregate: 'count' }],
				valuesAxis: 'rows',
				valuesPosition: 2,
				layout: 'outline',
				name: 'Many',
			},
		);
		const pivot = () => wb.sheets[1]!.pivotTables![0]!;
		const at = (position: number) => {
			edit.refreshPivotTable(1, 'Many', { valuesPosition: position });
			return pivotGrid(pivot()).values[1]!.slice(0, 4);
		};
		expect(at(2)).toEqual(['A', 'B', 'Values', 'C']);
		expect(at(0)).toEqual(['Values', 'A', 'B', 'C']);
		expect(at(3)).toEqual(['A', 'B', 'C', 'Values']);
		expect(() => at(5)).toThrow('Values before or between fields');
		edit.refreshPivotTable(1, 'Many', {
			fieldSubtotals: [
				{ field: 'A', functions: ['sum'] },
				{ field: 'D', functions: [] },
			],
		});
		expect(pivot().fieldSubtotals).toHaveLength(2);
	});
});
