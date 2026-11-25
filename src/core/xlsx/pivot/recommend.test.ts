import { expect, it } from 'vitest';
import { parseRange } from '../address';
import { createEditSession } from '../edit/session';
import { createWorkbook } from '../workbook';
import { loadXlsx } from '../read';
import { saveXlsx } from '../write';
import { recommendedPivotTables } from './index';

it('recommends grouped dates and categorical summaries and inserts the chosen full report', async () => {
	const book = createWorkbook({ sheets: ['Data', 'Report'] });
	const edit = createEditSession(book);
	[
		['Category', 'Date', 'Amount'],
		['Fruit', 45658, 10],
		['Vegetables', 45689, 20],
		['Fruit', 45700, 30],
	].forEach((row, r) => row.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
	edit.applyStyle(0, [parseRange('B2:B4')!], { numFmt: 'yyyy-mm-dd' });
	const recommendations = await recommendedPivotTables(book, 0, 'A1:C4', 'en');
	expect(recommendations.map(({ options }) => options.rowField)).toEqual(['Date', 'Category']);
	expect(recommendations[1]?.preview).toContainEqual(['Fruit', 40]);
	edit.createPivotTable(1, { row: 0, col: 0 }, recommendations[0]!.options);
	const reopened = await loadXlsx(await saveXlsx(book));
	const pivot = reopened.sheets[1]?.pivotTables?.[0];
	expect(pivot?.dateGroups?.[0]).toMatchObject({ field: 'Date', by: ['months', 'years'] });
	expect(pivot?.records).toHaveLength(3);
});

it('bounds preview rows and columns even for a much larger range', async () => {
	const book = createWorkbook();
	const edit = createEditSession(book);
	for (let c = 0; c < 32; c++) edit.setCellInput(0, 0, c, c === 0 ? 'Category' : `Value ${c}`);
	for (const row of [1, 9_999]) {
		edit.setCellInput(0, row, 0, 'Sample');
		for (let c = 1; c < 33; c++) edit.setCellInput(0, row, c, '2');
	}
	const result = await recommendedPivotTables(book, 0, 'A1:AG10000', 'en');
	expect(result).toHaveLength(4);
	expect(result[0]?.preview).toContainEqual(['Sample', 4]);
	expect(result.every(({ preview }) => preview.length <= 10)).toBe(true);
});
