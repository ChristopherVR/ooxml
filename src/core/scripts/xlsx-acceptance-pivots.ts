// Acceptance workbooks with PivotTables our writer authored: fixtures refreshed after loading,
// and reports built in an edit session (many row fields, custom subtotals, date groups). The
// PowerShell script opens each in Excel and refreshes every PivotTable, which must reproduce the
// cells we wrote. Used by `xlsx-acceptance-build.ts`.
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { createEditSession } from '../xlsx/edit/index';
import { ymdToSerial } from '../xlsx/formula/date-serial';
import type { Workbook } from '../xlsx/model';
import { loadXlsx } from '../xlsx/read/index';
import { createWorkbook } from '../xlsx/workbook';

const REFRESHED = [
	'rows3-compact-top',
	'rows3-compact-bottom',
	'rows3-outline-top-repeat',
	'rows3-outline-bottom',
	'rows3-tabular',
	'rows3-tabular-repeat',
	'rows3-tabular-columns',
	'subtotals-all',
	'subtotals-some-compact',
	'subtotals-none-compact',
	'subtotals-count-field',
	'group-date-months-years',
	'group-date-months-quarters-years-tabular',
	'group-date-days-7',
	'group-date-days-1',
	'group-date-time',
	'group-date-months-custom',
	'group-date-year-end',
	'cols2-compact',
	'cols2-tabular',
	'cols3-compact',
	'cols2-measures-compact',
	'cols2-measures-tabular',
	'cols2-rows2-compact-bottom',
	'cols2-rows2-subtotals',
];

async function refreshed(fixtures: string, name: string): Promise<Workbook> {
	const wb = await loadXlsx(readFileSync(join(fixtures, `excel-pivot-${name}.xlsx`)));
	const summary = wb.sheets.findIndex((sheet) => sheet.pivotTables?.length);
	createEditSession(wb).refreshPivotTable(summary, wb.sheets[summary]!.pivotTables![0]!.name);
	return wb;
}

/** Sales rows with a date column, a region and three nesting fields. */
function created(): Workbook {
	const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
	const s = createEditSession(wb);
	s.setRangeValues(0, { row: 0, col: 0 }, [
		['Date', 'Region', 'Product', 'Size', 'Amount'],
		...[
			[2024, 1, 5, 'East', 'Apple', 'S', 10],
			[2024, 2, 14, 'West', 'Apple', 'M', 20],
			[2024, 2, 28, 'East', 'Pear', 'S', 7.5],
			[2024, 4, 9, 'West', 'Pear', 'L', 2.25],
			[2024, 7, 30, 'East', 'Apple', 'M', 12],
			[2025, 1, 2, 'West', 'Plum', 'S', 3],
			[2025, 3, 3, 'East', 'Plum', 'L', 9],
			[2025, 11, 21, 'West', 'Apple', 'L', 15],
		].map(([y, m, d, ...rest]) => [
			ymdToSerial(y as number, m as number, d as number, false)!,
			...rest,
		]),
	] as never);
	s.applyStyle(0, [{ start: { row: 1, col: 0 }, end: { row: 8, col: 0 } }], {
		numFmt: 'yyyy-mm-dd',
	});
	const source = 'Data!A1:E9';
	s.createPivotTable(
		1,
		{ row: 0, col: 0 },
		{
			source,
			name: 'Nested',
			layout: 'outline',
			subtotalPosition: 'bottom',
			rowField: 'Region',
			secondRowField: 'Product',
			moreRowFields: ['Size'],
			valueField: 'Amount',
		},
	);
	s.createPivotTable(
		1,
		{ row: 0, col: 6 },
		{
			source,
			name: 'Subtotals',
			layout: 'compact',
			rowField: 'Region',
			secondRowField: 'Product',
			fieldSubtotals: [
				{ field: 'Region', functions: ['sum', 'average', 'max', 'min', 'stdDev', 'varp'] },
				{ field: 'Product', functions: [] },
			],
			valueField: 'Amount',
		},
	);
	s.createPivotTable(
		1,
		{ row: 0, col: 12 },
		{
			source,
			name: 'ByDate',
			layout: 'compact',
			rowField: 'Date',
			dateGroups: [{ field: 'Date', by: ['months', 'quarters', 'years'] }],
			valueField: 'Amount',
		},
	);
	s.createPivotTable(
		1,
		{ row: 0, col: 18 },
		{
			source,
			name: 'ByWeek',
			layout: 'tabular',
			rowField: 'Region',
			secondRowField: 'Date',
			dateGroups: [{ field: 'Date', by: ['days'], days: 30 }],
			valueField: 'Amount',
		},
	);
	s.createPivotTable(
		1,
		{ row: 30, col: 0 },
		{
			source,
			name: 'ColumnsNested',
			layout: 'tabular',
			rowField: 'Region',
			columnField: 'Product',
			moreColumnFields: ['Size'],
			valueField: 'Amount',
			additionalValues: [{ field: 'Amount', aggregate: 'count' }],
		},
	);
	s.createPivotTable(
		1,
		{ row: 50, col: 0 },
		{
			source,
			name: 'DateColumns',
			layout: 'compact',
			rowField: 'Region',
			columnField: 'Date',
			dateGroups: [{ field: 'Date', by: ['quarters', 'years'] }],
			valueField: 'Amount',
		},
	);
	return wb;
}

export const pivotBuilds = (fixtures: string): [string, () => Promise<Workbook> | Workbook][] => [
	...REFRESHED.map((name): [string, () => Promise<Workbook>] => [
		`pivot-refreshed-${name}.xlsx`,
		() => refreshed(fixtures, name),
	]),
	['pivot-created.xlsx', created],
];
