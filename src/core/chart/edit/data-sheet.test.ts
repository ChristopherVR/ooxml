import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { parseXml } from '../../xml/index';
import {
	applyChartSpaceDataSheetEdits,
	chartSpaceDataSheet,
	switchChartRowColumn,
} from './data-sheet';
import { DEFAULT_CHART_DATA_SHEET, newChartSpace } from './new-chart';
import { columnChart } from './test-support/column-chart';

const reload = (xml: string) => parseChartSpace(parseXml(xml).documentElement!).chartSpace;

describe('the data sheet of a chart part', () => {
	it('reads a new chart back as the sheet it was built from', () => {
		expect(chartSpaceDataSheet(newChartSpace())).toEqual(DEFAULT_CHART_DATA_SHEET);
	});

	it('keeps series formatting and a name reference through edits', () => {
		const base = newChartSpace();
		base.plotArea.groups[0]!.series[1]!.spPr = {
			fill: { kind: 'solid', color: { kind: 'srgb', value: 'FF0000', transforms: [] } },
		};
		const edited = applyChartSpaceDataSheetEdits(base, [
			{ kind: 'series-name', series: 1, name: 'Profit' },
			{ kind: 'remove-series', series: 0 },
			{ kind: 'remove-category', category: 3 },
		]);
		const [first] = edited.plotArea.groups[0]!.series;
		expect(first!.spPr?.fill).toMatchObject({ kind: 'solid' });
		expect(first!.tx?.reference?.formula).toBe('Sheet1!$B$1');
		expect(first!.values?.formula).toBe('Sheet1!$B$2:$B$4');
		expect(first!.categories?.formula).toBe('Sheet1!$A$2:$A$4');
		expect(chartSpaceDataSheet(edited)).toEqual({
			categories: ['Category 1', 'Category 2', 'Category 3'],
			series: [
				{ name: 'Profit', values: [2.4, 4.4, 1.8] },
				{ name: 'Series 3', values: [2, 2, 3] },
			],
		});
	});

	it('survives writing and reading the part', () => {
		const edited = applyChartSpaceDataSheetEdits(newChartSpace(), [
			{ kind: 'value', series: 2, category: 0, value: 12 },
			{ kind: 'add-series' },
		]);
		expect(chartSpaceDataSheet(reload(writeChartSpace(edited)))).toEqual(
			chartSpaceDataSheet(edited),
		);
	});

	it('sizes a series added after removing the longest one as a later save does', () => {
		const literal = (count: number) =>
			`<c:val><c:numLit><c:ptCount val="${count}"/>${Array.from(
				{ length: count },
				(_, index) => `<c:pt idx="${index}"><c:v>${index + 1}</c:v></c:pt>`,
			).join('')}</c:numLit></c:val>`;
		const series = (index: number, count: number) =>
			`<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${literal(count)}</c:ser>`;
		const ragged = reload(
			`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${series(0, 3)}${series(1, 8)}${series(2, 2)}</c:barChart></c:plotArea></c:chart></c:chartSpace>`,
		);
		const remove = { kind: 'remove-series', series: 1 } as const;
		const add = { kind: 'add-series' } as const;
		const once = applyChartSpaceDataSheetEdits(ragged, [remove, add]);
		const saved = reload(writeChartSpace(applyChartSpaceDataSheetEdits(ragged, [remove])));
		const twice = applyChartSpaceDataSheetEdits(saved, [add]);
		expect(chartSpaceDataSheet(once)).toEqual(chartSpaceDataSheet(twice));
		expect(chartSpaceDataSheet(once).series[2]!.values).toEqual([0, 0, 0]);
	});

	it('ignores edits the sheet refuses and returns the same model', () => {
		const space = newChartSpace();
		expect(
			applyChartSpaceDataSheetEdits(space, [{ kind: 'value', series: 9, category: 0, value: 1 }]),
		).toBe(space);
	});
});

describe('a new chart part', () => {
	it('switches to the picked family through the type switch', () => {
		const pie = newChartSpace(undefined, { kind: 'pie' }, { title: 'Sales' });
		expect(pie.plotArea.groups[0]!.element).toBe('pieChart');
		expect(pie.plotArea.axes).toEqual([]);
		expect(pie.title?.text).toBe('Sales');
		expect(pie.legend?.position).toBe('b');
		const bars = newChartSpace(undefined, {
			kind: 'bar',
			barDirection: 'bar',
			grouping: 'stacked',
		});
		expect(bars.plotArea.groups[0]).toMatchObject({ barDirection: 'bar', grouping: 'stacked' });
		expect(bars.autoTitleDeleted).toBe(true);
	});

	it('writes a part that reads back unchanged', () => {
		const line = newChartSpace(undefined, { kind: 'line' }, { title: 'Trend' });
		const xml = writeChartSpace(line);
		expect(xml).toContain('<c:lineChart>');
		expect(chartSpaceDataSheet(reload(xml))).toEqual(DEFAULT_CHART_DATA_SHEET);
	});
});

describe('the data sheet of a column chart part', () => {
	it('reads names, categories and values from the part caches', () => {
		expect(chartSpaceDataSheet(columnChart())).toEqual({
			categories: ['North', 'South', 'East'],
			series: [
				{ name: 'Plan', values: [10, 20, 30] },
				{ name: 'Actual', values: [12, 18, 33] },
			],
		});
	});

	it('applies edits through the series operations, keeping series formatting', () => {
		const space = applyChartSpaceDataSheetEdits(columnChart(), [
			{ kind: 'remove-category', category: 1 },
			{ kind: 'category-label', category: 0, label: 'N' },
			{ kind: 'remove-series', series: 0 },
			{ kind: 'value', series: 0, category: 1, value: 7 },
		]);
		expect(chartSpaceDataSheet(space)).toEqual({
			categories: ['N', 'East'],
			series: [{ name: 'Actual', values: [12, 7] }],
		});
		expect(space.plotArea.groups[0]!.series[0]!.spPr?.fill).toMatchObject({ kind: 'solid' });
	});

	it('returns the chart itself when the sheet refuses every edit', () => {
		const space = columnChart();
		expect(applyChartSpaceDataSheetEdits(space, [{ kind: 'remove-series', series: 5 }])).toBe(
			space,
		);
		const one = applyChartSpaceDataSheetEdits(space, [{ kind: 'remove-series', series: 0 }]);
		expect(applyChartSpaceDataSheetEdits(one, [{ kind: 'remove-series', series: 0 }])).toBe(one);
	});

	it('transposes categories and series on switchChartRowColumn and restores on double switch', () => {
		const orig = columnChart();
		const before = chartSpaceDataSheet(orig);
		const switched = switchChartRowColumn(orig);
		const mid = chartSpaceDataSheet(switched);
		expect(mid.categories).toEqual(before.series.map((s) => s.name));
		expect(mid.series.map((s) => s.name)).toEqual(before.categories);
		expect(mid.series[0]?.values).toEqual(before.series.map((s) => s.values[0]));

		const back = switchChartRowColumn(switched);
		expect(chartSpaceDataSheet(back)).toEqual(before);
	});
});
