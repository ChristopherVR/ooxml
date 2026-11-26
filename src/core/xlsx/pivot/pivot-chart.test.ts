import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { parseRange } from '../address';
import { createEditSession } from '../edit/session';
import type { PivotLayout, Workbook } from '../model';
import { loadXlsx } from '../read';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write';
import { chartView } from '../layout/chart-view';
import { pivotChartButtons, pivotChartTable } from './chart-binding';

const anchor = {
	from: { row: 0, col: 5, rowOffset: 0, colOffset: 0 },
	ext: { cx: 4_000_000, cy: 2_000_000 },
};

function chartOf(book: Workbook, index = 0) {
	const chart = book.sheets[2]?.drawings[index];
	if (chart?.kind !== 'chart') throw new Error('Missing chart');
	return chart;
}

function setup(table = false, layout: PivotLayout = 'tabular') {
	const book = createWorkbook({ sheets: ['Sales Data', 'Pivot Report', 'Charts'] });
	const edit = createEditSession(book);
	const rows = [
		['Region', 'Product', 'Amount', 'Year'],
		['North', 'Pens', 20, 2025],
		['South', 'Chairs', 15, 2026],
		['North', 'Chairs', 4, 2025],
	];
	rows.forEach((row, r) => row.forEach((value, c) => edit.setCellInput(0, r, c, String(value))));
	const source = table
		? edit.createTable(0, parseRange('A1:D4')!, true).name
		: "'Sales Data'!A1:D4";
	const name = edit.createPivotTable(
		1,
		{ row: 2, col: 1 },
		{
			source,
			rowField: 'Region',
			columnField: 'Product',
			valueField: 'Amount',
			layout,
			pageFilter: { field: 'Year' },
		},
	);
	edit.addPivotChart(2, 1, name, anchor);
	return { book, edit, name, source };
}

it.each([false, true])(
	'reads compact Excel pivot metadata when refresh has no required features (%s)',
	async (required) => {
		const { book } = setup(false, 'compact');
		const zip = await JSZip.loadAsync(await saveXlsx(book));
		const table = 'xl/pivotTables/pivotTable1.xml';
		const cache = 'xl/pivotCache/pivotCacheDefinition1.xml';
		const branch = required ? 'requiredFeature' : 'lastRefreshFeature';
		zip.file(
			table,
			(await zip.file(table)!.async('string')).replace(/<pivotTableDefinition[^>]*>/u, (tag) =>
				tag.replace(' outline="1" outlineData="1"', ''),
			),
		);
		zip.file(
			cache,
			(await zip.file(cache)!.async('string')).replace(
				'</pivotCacheDefinition>',
				`<extLst><ext uri="{9F748A41-CAEA-4470-BF7A-CE61E8FFA7F9}"><pvi:cacheVersionInfo xmlns:pvi="http://schemas.microsoft.com/office/spreadsheetml/2022/pivotVersionInfo"><pvi:${branch}>Autorefresh</pvi:${branch}></pvi:cacheVersionInfo></ext><ext uri="{F4336242-56EE-4DD3-A5F7-E900642BF6E9}"><pdc:dataSourceClean xmlns:pdc="http://schemas.microsoft.com/office/spreadsheetml/2026/pivotDataSourceClean">1</pdc:dataSourceClean></ext></extLst></pivotCacheDefinition>`,
			),
		);
		const reopened = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(pivotChartTable(reopened, chartOf(reopened))?.pivot.layout).toBe(
			required ? undefined : 'compact',
		);
	},
);

it.each([false, true])(
	'saves PivotChart parts and their workbook and table relationships (table source %s)',
	async (table) => {
		const { book, edit, name, source } = setup(table);
		edit.addPivotChart(2, 1, name, anchor);
		const bytes = await saveXlsx(book);
		const zip = await JSZip.loadAsync(bytes);
		const xml = (part: string) => zip.file(part)!.async('string');
		expect(await xml('xl/charts/chart1.xml')).toContain(
			'<c:pivotSource><c:name>[Book.xlsx]Pivot Report!PivotTable1</c:name><c:fmtId val="0"/>',
		);
		expect(await xml('xl/charts/chart2.xml')).toContain('<c:fmtId val="1"/>');
		expect(await xml('xl/charts/chart1.xml')).toContain(
			'<c:pivotFmt><c:idx val="1"/></c:pivotFmt>',
		);
		expect(await xml('xl/pivotTables/pivotTable1.xml')).toContain(
			'<chartFormat chart="1" format="1" series="1">',
		);
		expect(await xml('xl/pivotTables/_rels/pivotTable1.xml.rels')).toContain(
			'../pivotCache/pivotCacheDefinition1.xml',
		);
		expect(await xml('xl/_rels/workbook.xml.rels')).toContain(
			'pivotCache/pivotCacheDefinition1.xml',
		);
		expect(await xml('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels')).toContain(
			'pivotCacheRecords1.xml',
		);
		const reopened = await loadXlsx(bytes);
		const chart = chartOf(reopened);
		expect(pivotChartTable(reopened, chart)).toMatchObject({
			sheet: 1,
			pivot: { name, source, records: expect.any(Array) },
		});
		expect(chart.series.map(({ name, values }) => [name, values])).toEqual([
			['Chairs', [4, 15]],
			['Pens', [20, null]],
		]);
		expect(pivotChartButtons(chart)).toEqual(new Set(['axis', 'legend', 'filter', 'value']));
	},
);

it('refreshes cross-sheet chart caches after field, source and filter edits, including undo and reopen', async () => {
	const { book, edit, name } = setup();
	edit.setCellInput(0, 1, 2, '30');
	edit.refreshPivotTable(1, name);
	expect(chartOf(book).series[1]?.values).toEqual([30, null]);
	edit.refreshPivotTable(1, name, {
		pageFilter: { field: 'Year', item: 2025 },
		itemFilters: [{ field: 'Product', hidden: ['Chairs'] }],
	});
	expect(chartOf(book).series.map((series) => series.values)).toEqual([[30]]);
	edit.undo();
	expect(chartOf(book).series.map((series) => series.name)).toEqual(['Chairs', 'Pens']);
	edit.redo();
	edit.refreshPivotTable(1, name, { aggregate: 'count', columnField: '' });
	expect(chartOf(book).series[0]).toMatchObject({
		name: 'Count of Amount',
		values: [1],
		categories: ['North'],
	});
	const view = chartView(book, 2, chartOf(book), () => ['incorrect chart reference evaluation']);
	expect(view.categories).toEqual(['North']);
	edit.showPivotChartButtons(2, 0, false);
	const reopened = await loadXlsx(await saveXlsx(book));
	expect(pivotChartButtons(chartOf(reopened)).size).toBe(0);
	expect(chartOf(reopened).series[0]?.values).toEqual([1]);
	createEditSession(reopened).refreshPivotTable(1, name, { pageFilter: null, itemFilters: [] });
	expect(chartOf(reopened).series[0]?.values).toEqual([2, 1]);
});

it('tracks report and source names and detaches the chart with its cached data on deletion', async () => {
	const { book, edit, name, source } = setup(true);
	edit.updateTable(0, source, { name: 'NewSource' });
	edit.renameSheet(1, 'New Report');
	const saved = await loadXlsx(await saveXlsx(book));
	expect(pivotChartTable(saved, chartOf(saved))?.pivot.source).toBe('NewSource');
	const reopened = createEditSession(saved);
	reopened.convertTableToRange(0, 'NewSource');
	reopened.refreshPivotTable(1, name);
	reopened.deletePivotTable(1, name);
	expect(chartOf(saved).series[0]?.values).toEqual([4, 15]);
	expect(chartOf(saved).series[0]?.valuesRef).toBeUndefined();
	reopened.undo();
	expect(pivotChartTable(saved, chartOf(saved))?.pivot.name).toBe(name);
	reopened.redo();
	const detached = chartOf(await loadXlsx(await saveXlsx(saved)));
	expect(detached.chartSpace?.pivotSource).toBeUndefined();
	expect(detached.series[0]?.values).toEqual([4, 15]);
});

it('keeps a loaded PivotChart workbook and sheet prefix when renaming its report', async () => {
	const { book, edit, name } = setup();
	edit.renameSheet(1, 'Pivot!Report');
	const part = 'xl/charts/chart1.xml';
	const zip = await JSZip.loadAsync(await saveXlsx(book));
	const xml = await zip.file(part)!.async('string');
	zip.file(part, xml.replace('[Book.xlsx]', '[RealName.xlsx]'));

	const loaded = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	const session = createEditSession(loaded);
	session.updatePivotTable(1, name, { name: 'RenamedPivot' });
	expect(chartOf(loaded).chartSpace?.pivotSource?.name).toBe(
		'[RealName.xlsx]Pivot!Report!RenamedPivot',
	);
	session.undo();
	expect(chartOf(loaded).chartSpace?.pivotSource?.name).toBe(`[RealName.xlsx]Pivot!Report!${name}`);
	session.redo();
	const reopened = await loadXlsx(await saveXlsx(loaded));
	expect(chartOf(reopened).chartSpace?.pivotSource?.name).toBe(
		'[RealName.xlsx]Pivot!Report!RenamedPivot',
	);
	expect(pivotChartTable(reopened, chartOf(reopened))?.pivot.name).toBe('RenamedPivot');
});
