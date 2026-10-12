import { describe, expect, it } from 'vitest';
import type { ChartObject } from '../model';
import { createEditSession } from '../edit/session';
import { parseChart } from '../read/chart';
import { createWorkbook } from '../workbook';
import { chartXml } from './chart';
import { chartWithCurrentCaches } from './chart-cache';

/** A chart with stale caches over live cells, including blanks and non-numeric values. */
function setup() {
	const workbook = createWorkbook({ sheets: ['Data'] });
	const session = createEditSession(workbook, { autoRowHeight: false });
	session.setRangeValues(0, { row: 0, col: 0 }, [
		['Current', 2.5, 10],
		[null, null, null],
		[null, { error: '#N/A' }, 'text'],
	]);
	const chart: ChartObject = {
		kind: 'chart',
		anchor: { from: { row: 4, col: 0, rowOffset: 0, colOffset: 0 } },
		chartType: 'column',
		showLegend: true,
		series: [
			{
				name: 'cached',
				nameRef: 'data!A1',
				categoriesRef: 'data!B1:B3',
				categories: ['old'],
				valuesRef: 'C1:C3',
				values: [123],
			},
		],
	};
	return { workbook, session, chart, sheet: workbook.sheets[0]! };
}

describe('current caches for regenerated charts', () => {
	it('saves current values and blank positions without mutating the chart', () => {
		const { workbook, chart, sheet } = setup();
		const before = structuredClone(chart);
		const current = chartWithCurrentCaches(workbook, sheet, chart);
		expect(current.series[0]?.categories).toEqual([2.5, '', '#N/A']);
		const saved = parseChart(chartXml(current), chart.anchor, 'xl/charts/chart1.xml');
		expect(saved.series[0]).toMatchObject({
			name: 'Current',
			categories: ['2.5', '', '#N/A'],
			values: [10, null, null],
		});
		expect(chart).toEqual(before);
	});

	it.each(['data!A1:A3', 'data!A1:B1'])(
		'keeps the cached series name for a multi-cell name reference: %s',
		(ref) => {
			const { workbook, chart, sheet } = setup();
			chart.series[0]!.nameRef = ref;
			const before = structuredClone(chart);
			const current = chartWithCurrentCaches(workbook, sheet, chart);
			const saved = parseChart(chartXml(current), chart.anchor, 'xl/charts/chart1.xml');
			expect(saved.series[0]).toMatchObject({
				name: 'cached',
				nameRef: ref,
				categories: ['2.5', '', '#N/A'],
				values: [10, null, null],
			});
			expect(chart).toEqual(before);
		},
	);

	it.each([
		{ value: 0, expected: '0' },
		{ value: false, expected: 'FALSE' },
		{ value: null, expected: '' },
		{ value: '', expected: '' },
	])('refreshes a single-cell name without falling back for %j', ({ value, expected }) => {
		const { workbook, session, chart, sheet } = setup();
		session.setCellValue(0, 0, 0, value);
		const current = chartWithCurrentCaches(workbook, sheet, chart);
		const saved = parseChart(chartXml(current), chart.anchor, 'xl/charts/chart1.xml');
		expect(saved.series[0]?.name).toBe(expected);
	});

	it('keeps the name cache for a whole-column name reference', () => {
		const { workbook, chart, sheet } = setup();
		chart.series[0]!.nameRef = 'data!A:A';
		const current = chartWithCurrentCaches(workbook, sheet, chart);
		expect(current.series[0]?.name).toBe('cached');
	});

	it('keeps whole-column caches without reading their cells', () => {
		const { workbook, chart, sheet } = setup();
		const series = chart.series[0]!;
		series.valuesRef = 'data!C1:C1048576';
		series.categoriesRef = 'data!B1:B1048576';
		for (const col of [1, 2])
			Object.defineProperty(sheet.rows.get(0)!.get(col)!, 'value', {
				get() {
					throw new Error('Oversized chart reference must not read cells');
				},
			});
		const current = chartWithCurrentCaches(workbook, sheet, chart);
		expect(current.series[0]?.values).toBe(series.values);
		expect(current.series[0]?.categories).toBe(series.categories);
		const saved = parseChart(chartXml(current), chart.anchor, 'xl/charts/chart1.xml');
		expect(saved.series[0]).toMatchObject({
			valuesRef: series.valuesRef,
			values: [123],
			categoriesRef: series.categoriesRef,
			categories: ['old'],
		});
	});

	it.each(['values', 'categories'] as const)(
		'keeps the complete %s cache above the chart point limit',
		(kind) => {
			const { workbook, chart, sheet } = setup();
			const series = chart.series[0]!;
			if (kind === 'values') series.valuesRef = 'data!C1:C10001';
			else series.categoriesRef = 'data!B1:B10001';
			const current = chartWithCurrentCaches(workbook, sheet, chart);
			expect(current.series[0]?.[kind]).toBe(series[kind]);
			expect(current.series[0]?.name).toBe('Current');
		},
	);

	it('refreshes all points at the chart cache limit', () => {
		const { workbook, session, chart, sheet } = setup();
		session.setCellValue(0, 9999, 2, 42);
		chart.series[0]!.valuesRef = 'data!C1:C10000';
		const current = chartWithCurrentCaches(workbook, sheet, chart);
		const saved = parseChart(chartXml(current), chart.anchor, 'xl/charts/chart1.xml');
		expect(saved.series[0]?.values).toHaveLength(10_000);
		expect(saved.series[0]?.values[0]).toBe(10);
		expect(saved.series[0]?.values[9999]).toBe(42);
	});

	it('keeps numeric scatter categories numeric', () => {
		const { workbook, session, chart, sheet } = setup();
		session.setRangeValues(0, { row: 0, col: 1 }, [[1], [2.5], [3]]);
		chart.chartType = 'scatter';
		const current = chartWithCurrentCaches(workbook, sheet, chart);
		const saved = parseChart(chartXml(current), chart.anchor, 'xl/charts/chart1.xml');
		expect(saved.series[0]?.categories).toEqual([1, 2.5, 3]);
	});

	it.each([
		'Missing!A1:A2',
		"'[remote.xlsx]Data'!A1:A2",
		'Data!NamedInput',
		'Data!A:XFD',
		'Data!A1048577',
	])('keeps caches when %s cannot be read safely', (ref) => {
		const { workbook, chart, sheet } = setup();
		Object.assign(chart.series[0]!, { nameRef: ref, categoriesRef: ref, valuesRef: ref });
		const before = structuredClone(chart);
		expect(chartWithCurrentCaches(workbook, sheet, chart)).toEqual(before);
		expect(chart).toEqual(before);
	});

	it('keeps literal series without references', () => {
		const { workbook, chart, sheet } = setup();
		chart.series = [{ name: 'literal', categories: ['a'], values: [5] }];
		expect(chartWithCurrentCaches(workbook, sheet, chart)).toEqual(chart);
	});
});
