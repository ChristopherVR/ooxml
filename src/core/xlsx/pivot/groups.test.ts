import { readFileSync, writeFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { getCell } from '../cells';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/index';
import { defaultFieldGroup, numberGroupLabel, numberGroupLabels } from './groups';

const fixture = readFileSync(
	new URL('../__fixtures__/excel-pivot-group-number.xlsx', import.meta.url),
);
const text = async (bytes: Uint8Array, name: string) =>
	(await (await JSZip.loadAsync(bytes)).file(name)?.async('string')) ?? '';
const column = (sheet: Parameters<typeof getCell>[0], rows: number[]) =>
	rows.map((row) => [getCell(sheet, row, 0)?.value, getCell(sheet, row, 1)?.value]);

it('labels number groups like Excel and assigns values to buckets', () => {
	const group = { field: 'Score', start: 0, end: 50, by: 10 };
	expect(numberGroupLabels(group)).toEqual([
		'<0',
		'0-9',
		'10-19',
		'20-29',
		'30-39',
		'40-50',
		'>50',
	]);
	expect(numberGroupLabel(group, 9)).toBe('0-9');
	expect(numberGroupLabel(group, 10)).toBe('10-19');
	expect(numberGroupLabel(group, 50)).toBe('40-50');
	expect(numberGroupLabel(group, 51)).toBe('>50');
	expect(numberGroupLabel(group, -1)).toBe('<0');
	// An end that is not a multiple of the interval closes the last bucket early.
	expect(numberGroupLabels({ start: 1, end: 25, by: 10 })).toEqual([
		'<1',
		'1-10',
		'11-20',
		'21-25',
		'>25',
	]);
});

it('reads a native number group, refreshes it and round-trips unchanged parts', async () => {
	const wb = await loadXlsx(fixture);
	expect(wb.warnings).toEqual([]);
	const sheet = wb.sheets[0]!;
	const pivot = sheet.pivotTables![0]!;
	expect(pivot.fieldGroups).toEqual([{ field: 'Score', start: 0, end: 50, by: 10 }]);
	expect(column(sheet, [3, 4, 5, 6, 7])).toEqual([
		['0-9', 10],
		['10-19', 25],
		['20-29', 22],
		['40-50', 9],
		['Grand Total', 66],
	]);
	const original = await JSZip.loadAsync(fixture);
	const saved = await JSZip.loadAsync(await saveXlsx(wb));
	for (const name of Object.keys(original.files).filter((n) => n.includes('xl/pivot')))
		expect(await saved.file(name)?.async('string'), name).toBe(
			await original.file(name)!.async('string'),
		);
	const edit = createEditSession(wb);
	edit.refreshPivotTable(0, 'GroupPivot');
	const rewritten = await saveXlsx(wb);
	const cache = await text(rewritten, 'xl/pivotCache/pivotCacheDefinition2.xml');
	expect(cache).toContain(
		'<fieldGroup base="1"><rangePr autoStart="0" autoEnd="0" startNum="0" endNum="50" groupInterval="10"/>',
	);
	expect(cache).toContain('<groupItems count="7"><s v="&lt;0"/><s v="0-9"/>');
	const table = await text(rewritten, 'xl/pivotTables/pivotTable1.xml');
	expect(table).toContain('<items count="8"><item x="0"/><item x="1"/>');
	expect(table).toContain(
		'<rowItems count="5"><i><x v="1"/></i><i><x v="2"/></i><i><x v="3"/></i><i><x v="5"/></i>',
	);
	const again = await loadXlsx(rewritten);
	expect(again.warnings).toEqual([]);
	expect(again.sheets[0]!.pivotTables![0]!.fieldGroups).toEqual(pivot.fieldGroups);
});

it('groups and ungroups a numeric field and keeps the groups through a source refresh', async () => {
	const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
	const edit = createEditSession(wb);
	const rows = [
		['Category', 'Score', 'Amount'],
		['A', '3', '10'],
		['B', '12', '20'],
		['A', '17', '5'],
		['B', '25', '15'],
	];
	rows.forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, value)));
	edit.createPivotTable(
		1,
		{ row: 0, col: 0 },
		{ source: 'Data!A1:C5', rowField: 'Score', valueField: 'Amount' },
	);
	const summary = wb.sheets[1]!;
	expect(column(summary, [1, 2, 3, 4])).toEqual([
		[3, 10],
		[12, 20],
		[17, 5],
		[25, 15],
	]);
	edit.refreshPivotTable(1, 'PivotTable1', {
		fieldGroups: [{ field: 'Score', start: 0, end: 30, by: 10 }],
	});
	expect(column(summary, [1, 2, 3, 4])).toEqual([
		['0-9', 10],
		['10-19', 25],
		['20-30', 15],
		['Grand Total', 50],
	]);
	edit.setCellInput(0, 1, 1, '5');
	edit.refreshPivotTable(1, 'PivotTable1');
	expect(column(summary, [1])).toEqual([['0-9', 10]]);
	const saved = await saveXlsx(wb);
	if (process.env.OOXML_DUMP) writeFileSync(process.env.OOXML_DUMP, saved);
	const again = await loadXlsx(saved);
	expect(again.warnings).toEqual([]);
	expect(again.sheets[1]!.pivotTables![0]!.fieldGroups).toHaveLength(1);
	edit.refreshPivotTable(1, 'PivotTable1', { fieldGroups: [] });
	expect(column(summary, [1])).toEqual([[5, 10]]);
	expect(() =>
		edit.refreshPivotTable(1, 'PivotTable1', {
			fieldGroups: [{ field: 'Category', start: 0, end: 30, by: 10 }],
		}),
	).toThrow('Only fields that contain numbers can be grouped.');
	expect(() =>
		edit.refreshPivotTable(1, 'PivotTable1', {
			fieldGroups: [{ field: 'Score', start: 0, end: 30, by: 0 }],
		}),
	).toThrow('Choose a valid number group.');
});

it('offers a group over more records than a call takes arguments', () => {
	// About 125,000 is the most arguments a call takes in Node; the extremes sit mid-column.
	const score = (row: number) =>
		row === 70_000 ? -5 : row === 130_000 ? 2_500 : 100 + (row % 100);
	const records = Array.from({ length: 200_000 }, (_, row) => ['Fruit', score(row)]);
	expect(defaultFieldGroup({ fields: ['Category', 'Score'], records }, 'Score')).toEqual({
		field: 'Score',
		start: -5,
		end: 2_500,
		by: 251,
	});
});
