// Acceptance workbooks with PivotTables our writer authored: fixtures refreshed after loading,
// and reports built in an edit session (many row fields, custom subtotals, date groups). The
// PowerShell script opens each in Excel and refreshes every PivotTable, which must reproduce the
// cells we wrote. Used by `xlsx-acceptance-build.ts`.
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { createEditSession } from '../xlsx/edit/index';
import { ymdToSerial } from '../xlsx/formula/date-serial';
import type { PivotTableOptions, Workbook } from '../xlsx/model';
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
	// Custom subtotals on column fields.
	'colsubs-none-compact',
	'colsubs-sum-average-compact',
	'colsubs-sum-average-tabular',
	'colsubs-count-outline',
	'colsubs-three-levels',
	'colsubs-measures-compact',
	'colsubs-measures-other',
	'colsubs-rows-subtotals',
	'colsubs-rows-custom',
	'colsubs-rows-custom-compact',
	'colsubs-rows-same',
	'colsubs-rows-total',
	'colsubs-rows-auto',
	// Values on rows with any number of row and column fields, and subtotals on them.
	'rv3-last-compact-top',
	'rv3-last-outline-bottom',
	'rv3-last-outline-top-repeat',
	'rv3-last-tabular',
	'rv3-first-compact-top',
	'rv3-first-compact-bottom',
	'rv3-first-outline-top',
	'rv3-first-tabular-repeat',
	'rv3-second-compact-top',
	'rv3-second-outline-top-repeat',
	'rv3-second-tabular',
	'rv3-third-outline-top',
	'rv3-third-tabular-repeat',
	'rv2-cols2-last-compact',
	'rv2-cols2-first-tabular',
	'rv2-cols2-second-outline',
	'rv1-cols2-compact',
	'rv1-first-compact',
	'rv1-first-outline-cols',
	'rvs-last-compact',
	'rvs-first-compact-top',
	'rvs-first-compact-bottom',
	'rvs-second-tabular',
	'rvs-third-outline',
	'rvs-cols-compact',
	'rvs-rows-cols-tabular',
	'rvs-first-tabular-cols',
	'rvs-none-last',
	'rvs-none-first',
	// Values on rows under a report filter.
	'rvp-last-compact',
	'rvp-first-tabular',
	'rvp-second-outline',
	'rvp-cols-compact',
	'rvp-tabular-repeat',
	'rvp-first-compact-bottom',
	// Item display names written by Excel (PivotItem.Caption).
	'captions-items-tabular',
	'captions-items-compact-page',
	// Date groups with report filters, hidden items, a slicer and a timeline.
	'group-page-all',
	'group-page-west',
	'group-page-tabular-quarters',
	'group-page-columns',
	'group-hidden-items',
	'group-slicer',
	'group-timeline',
	'group-timeline-days',
	'group-page-timeline',
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
	s.createPivotTable(
		1,
		{ row: 70, col: 0 },
		{
			source,
			name: 'ColumnSubtotals',
			layout: 'tabular',
			rowField: 'Region',
			columnField: 'Product',
			moreColumnFields: ['Size'],
			fieldSubtotals: [
				{ field: 'Product', functions: ['sum', 'average', 'max'] },
				{ field: 'Size', functions: [] },
			],
			valueField: 'Amount',
			additionalValues: [{ field: 'Amount', aggregate: 'count' }],
		},
	);
	s.createPivotTable(
		1,
		{ row: 90, col: 0 },
		{
			source,
			name: 'RowValues',
			layout: 'outline',
			rowField: 'Region',
			secondRowField: 'Product',
			moreRowFields: ['Size'],
			valuesAxis: 'rows',
			valuesPosition: 1,
			fieldSubtotals: [{ field: 'Region', functions: ['sum', 'average'] }],
			valueField: 'Amount',
			additionalValues: [{ field: 'Amount', aggregate: 'count' }],
		},
	);
	s.createPivotTable(
		1,
		{ row: 150, col: 0 },
		{
			source,
			name: 'RowValuesFirst',
			layout: 'compact',
			rowField: 'Region',
			secondRowField: 'Product',
			valuesAxis: 'rows',
			valuesPosition: 'first',
			columnField: 'Size',
			valueField: 'Amount',
			additionalValues: [{ field: 'Amount', aggregate: 'max' }],
		},
	);
	s.createPivotTable(
		1,
		{ row: 190, col: 0 },
		{
			source,
			name: 'DatesWithPage',
			layout: 'compact',
			rowField: 'Date',
			dateGroups: [{ field: 'Date', by: ['months', 'years'] }],
			pageFilter: { field: 'Region', item: 'East' },
			valueField: 'Amount',
		},
	);
	// Values on rows under a report filter, with the Values among, last and before the row fields.
	const rowValuesBase: PivotTableOptions = {
		source,
		rowField: 'Product',
		valuesAxis: 'rows',
		valueField: 'Amount',
		additionalValues: [{ field: 'Amount', aggregate: 'count' }],
	};
	const underFilter: [string, PivotTableOptions][] = [
		[
			'compact',
			{
				...rowValuesBase,
				layout: 'compact',
				secondRowField: 'Size',
				valuesPosition: 'between',
				pageFilter: { field: 'Region', item: 'East' },
			},
		],
		[
			'outline',
			{ ...rowValuesBase, layout: 'outline', columnField: 'Size', pageFilter: { field: 'Region' } },
		],
		[
			'tabular',
			{
				...rowValuesBase,
				layout: 'tabular',
				secondRowField: 'Size',
				valuesPosition: 'first',
				pageFilter: { field: 'Region', item: 'West' },
			},
		],
	];
	// Captions Excel keeps in the definition (header, grand total, Values) and the value names. The
	// words only a UI language supplies stay English: an English Excel refreshes them as English.
	s.createPivotTable(
		1,
		{ row: 320, col: 0 },
		{
			source,
			name: 'GermanCaptions',
			layout: 'compact',
			rowField: 'Product',
			columnField: 'Size',
			valueField: 'Amount',
			additionalValues: [{ field: 'Amount', aggregate: 'count', caption: 'Anzahl von Amount' }],
			captions: {
				rowLabels: 'Zeilenbeschriftungen',
				columnLabels: 'Spaltenbeschriftungen',
				grandTotal: 'Gesamtergebnis',
				values: 'Werte',
			},
		},
	);
	underFilter.forEach(([layout, options], i) =>
		s.createPivotTable(
			1,
			{ row: 210 + i * 30, col: 0 },
			{ ...options, name: `RowValuesPage${layout}` },
		),
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
