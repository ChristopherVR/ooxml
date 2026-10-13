import { readFileSync } from 'node:fs';
import { expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createEditSession } from '../edit/session';
import { pivotGrid } from './aggregate';
import * as measures from './measures';
import type { PivotTable } from '../model';
import { parseXml, elements } from '../../xml/index';
import { supportedPivotAxes } from './values-axis';

for (const suffix of ['', '-columns'])
	it(`imports native Values last on rows ${suffix}, refreshes and preserves source parts`, async () => {
		const bytes = readFileSync(
			new URL(`../__fixtures__/excel-pivot-row-values${suffix}.xlsx`, import.meta.url),
		);
		const wb = await loadXlsx(bytes);
		expect(wb.warnings).toEqual([]);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(pivot.valuesAxis).toBe('rows');
		expect(pivot.secondRowField).toBeUndefined();
		expect(
			pivotGrid(pivot)
				.values.slice(-3)
				.map((row) => row.at(-1)),
		).toEqual([57, 5, 11.4]);
		const source = await JSZip.loadAsync(bytes);
		const untouched = await JSZip.loadAsync(await saveXlsx(wb));
		for (const part of Object.keys(source.files).filter((part) => part.startsWith('xl/pivot')))
			expect(await untouched.file(part)?.async('string')).toBe(
				await source.file(part)!.async('string'),
			);
		const edit = createEditSession(wb);
		edit.setCellInput(1, 1, 2, '100');
		edit.refreshPivotTable(0, pivot.name);
		const loaded = await loadXlsx(await saveXlsx(wb));
		expect(loaded.warnings).toEqual([]);
		expect(
			pivotGrid(loaded.sheets[0]!.pivotTables![0]!)
				.values.slice(-3)
				.map((row) => row.at(-1)),
		).toEqual([147, 5, 29.4]);
		edit.undo();
		expect(
			pivotGrid(wb.sheets[0]!.pivotTables![0]!)
				.values.slice(-3)
				.map((row) => row.at(-1)),
		).toEqual([57, 5, 11.4]);
	});

it('switches Values placement atomically and requires a supported profile', async () => {
	const wb = await loadXlsx(
		readFileSync(new URL('../__fixtures__/excel-pivot-row-values-columns.xlsx', import.meta.url)),
	);
	const edit = createEditSession(wb);
	const current = () => wb.sheets[0]!.pivotTables![0]!;
	edit.refreshPivotTable(0, current().name, { layout: 'compact' });
	expect(current().layout).toBe('compact');
	edit.undo();
	expect(() =>
		edit.refreshPivotTable(0, current().name, { secondRowField: 'Region', layout: 'compact' }),
	).toThrow('tabular');
	expect(() => edit.refreshPivotTable(0, current().name, { additionalValues: [] })).toThrow(
		'multiple',
	);
	edit.refreshPivotTable(0, current().name, { valuesAxis: 'columns' });
	expect(pivotGrid(current()).values.at(-1)?.slice(-3)).toEqual([57, 5, 11.4]);
	const loaded = await loadXlsx(await saveXlsx(wb));
	expect(loaded.warnings).toEqual([]);
	expect(loaded.sheets[0]!.pivotTables![0]?.valuesAxis).not.toBe('rows');
	edit.undo();
	expect(current().valuesAxis).toBe('rows');
});

it.each(['Values first', 'conflicting axis flag', 'bad row measure index'])(
	'preserves unsupported %s',
	async (kind) => {
		const zip = await JSZip.loadAsync(
			readFileSync(new URL('../__fixtures__/excel-pivot-row-values.xlsx', import.meta.url)),
		);
		const part = 'xl/pivotTables/pivotTable1.xml';
		const source = await zip.file(part)!.async('string');
		const xml =
			kind === 'Values first'
				? source.replace('<field x="0"/><field x="-2"/>', '<field x="-2"/><field x="0"/>')
				: kind === 'conflicting axis flag'
					? source.replace('dataOnRows="1"', 'dataOnRows="0"')
					: source.replace('i="1"', 'i="99"');
		expect(xml).not.toBe(source);
		zip.file(part, xml);
		const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(wb.sheets[0]!.pivotTables).toBeUndefined();
		expect(wb.warnings.some((warning) => warning.includes('cannot be refreshed'))).toBe(true);
		expect(await (await JSZip.loadAsync(await saveXlsx(wb))).file(part)!.async('string')).toBe(xml);
	},
);

it('rejects a large Values row result before allocating headers or measure bands', () => {
	const pivot: PivotTable = {
		name: 'Large',
		source: 'A1:C10000',
		fields: ['Group', 'Column', 'Amount'],
		rowField: 'Group',
		columnField: 'Column',
		valueField: 'Amount',
		valueCaption: 'Sum of Amount',
		aggregate: 'sum',
		valuesAxis: 'rows',
		layout: 'tabular',
		additionalValues: Array.from({ length: 15 }, (_, i) => ({
			field: 'Amount',
			aggregate: 'sum',
			caption: `Measure ${i}`,
		})),
		location: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
		records: Array.from({ length: 21000 }, (_, i) => [`Group ${i % 2}`, i, 1]),
	};
	const headers = vi.spyOn(measures, 'pivotHeaders');
	const numbers = vi.spyOn(measures, 'pivotNumbers');
	try {
		expect(() => pivotGrid(pivot)).toThrow('too large');
		expect(headers).not.toHaveBeenCalled();
		expect(numbers).not.toHaveBeenCalled();
	} finally {
		headers.mockRestore();
		numbers.mockRestore();
	}
});

it('converts native row measures back to one measure without stale Values metadata', async () => {
	const wb = await loadXlsx(
		readFileSync(new URL('../__fixtures__/excel-pivot-row-values-columns.xlsx', import.meta.url)),
	);
	const edit = createEditSession(wb);
	const pivot = wb.sheets[0]!.pivotTables![0]!;
	edit.refreshPivotTable(0, pivot.name, { valuesAxis: 'columns', additionalValues: [] });
	const bytes = await saveXlsx(wb);
	const zip = await JSZip.loadAsync(bytes);
	const xml = await zip.file('xl/pivotTables/pivotTable1.xml')!.async('string');
	expect(xml).toContain('dataOnRows="0"');
	expect(xml).not.toContain('x="-2"');
	const loaded = await loadXlsx(bytes);
	expect(loaded.warnings).toEqual([]);
	expect(pivotGrid(loaded.sheets[0]!.pivotTables![0]!).values.at(-1)).toEqual([
		'Grand Total',
		15,
		42,
		57,
	]);
	edit.undo();
	expect(wb.sheets[0]!.pivotTables![0]?.valuesAxis).toBe('rows');
});

it.each(['', 'other', false])(
	'refuses an invalid public Values axis %s without changing cells or history',
	async (axis) => {
		const wb = await loadXlsx(
			readFileSync(new URL('../__fixtures__/excel-pivot-row-values.xlsx', import.meta.url)),
		);
		const edit = createEditSession(wb);
		const pivot = wb.sheets[0]!.pivotTables![0]!;
		expect(() =>
			edit.createPivotTable(
				0,
				{ row: 20, col: 0 },
				{
					source: pivot.source,
					rowField: pivot.rowField,
					valueField: pivot.valueField,
					valuesAxis: axis as 'rows',
				},
			),
		).toThrow('valid Values axis');
		expect(() => edit.refreshPivotTable(0, pivot.name, { valuesAxis: axis as 'rows' })).toThrow(
			'valid Values axis',
		);
		expect(wb.sheets[0]!.pivotTables![0]).toBe(pivot);
		expect(edit.canUndo()).toBe(false);
	},
);

it.each(['', 'x=""', 'x="0.5"', 'x="NaN"'])(
	'refuses malformed Values axis field index %s',
	(attribute) => {
		const fields = elements(
			parseXml(`<fields><field ${attribute}/><field x="-2"/></fields>`).documentElement,
		);
		expect(supportedPivotAxes(fields, [], 2, true)).toBe(false);
		expect(supportedPivotAxes(fields, [], 2, false)).toBe(false);
	},
);
