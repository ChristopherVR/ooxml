import { describe, expect, it } from 'vitest';
import { columnChart } from './test-support/column-chart';
import { joinChartGroups, setChartSeriesType } from './series-type';
import { setChartSeriesSecondaryAxis } from './secondary-axis';
import { setChartType } from './type';
import { chartSeriesEntries } from './draft';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { chartSummaryFromSpace } from '../render/summary-from-space';
import { renderChartSpaceSvg } from '../render/render-space';

describe('one series combo type changes', () => {
	it('uses the shared family conversion while retaining every document-order identity and axis', () => {
		const source = columnChart(),
			original = structuredClone(source);
		const changed = setChartSeriesType(source, 1, { kind: 'line', is3D: false });
		expect(source).toEqual(original);
		const saved = parseChartSpace(writeChartSpace(changed)).chartSpace;
		expect(saved.plotArea.groups.map((group) => group.kind)).toEqual(['bar', 'line']);
		expect(saved.plotArea.groups.map((group) => group.axisIds)).toEqual([
			[11, 22],
			[11, 22],
		]);
		expect(saved.plotArea.axes).toEqual(source.plotArea.axes);
		expect(chartSeriesEntries(saved).map((entry) => entry.series.index)).toEqual([0, 1]);
		expect(chartSeriesEntries(saved).map((entry) => entry.series.values)).toEqual(
			chartSeriesEntries(source).map((entry) => entry.series.values),
		);
		expect(chartSeriesEntries(saved)[0]?.series).toEqual(chartSeriesEntries(source)[0]?.series);
		expect(chartSeriesEntries(saved)[1]?.series.spPr?.line?.fill).toEqual(
			chartSeriesEntries(source)[1]?.series.spPr?.fill,
		);
		expect(chartSummaryFromSpace(saved).extraGroups?.[0]?.chartType).toBe('line');
		expect(renderChartSpaceSvg(saved, { width: 480, height: 300 })).toContain('<polyline');
	});
	it('supports the first series without moving subsequent selections and refuses unsupported families', () => {
		const source = columnChart();
		const changed = setChartSeriesType(source, 0, { kind: 'area' });
		expect(changed.plotArea.groups.map((group) => group.kind)).toEqual(['area', 'bar']);
		expect(chartSeriesEntries(changed).map((entry) => entry.series.index)).toEqual([0, 1]);
		expect(() => setChartSeriesType(source, 1, { kind: 'bubble' })).toThrow('vertical 2D');
		expect(() => setChartSeriesType(source, 1, { kind: 'pie' })).toThrow('vertical 2D');
		expect(() => setChartSeriesType(source, 1, { kind: 'line', is3D: true })).toThrow(
			'vertical 2D',
		);
		source.plotArea.groups[0]!.extLst = '<c:extLst/>';
		expect(() => setChartSeriesType(source, 1, { kind: 'line' })).toThrow('unknown group');
	});
	it('writes a scatter series on the category and value axes it leaves, as Excel saves it', () => {
		const source = columnChart();
		const changed = setChartSeriesType(source, 1, { kind: 'scatter' });
		const saved = parseChartSpace(writeChartSpace(changed));
		expect(saved.issues).toEqual([]);
		const [bar, scatter] = saved.chartSpace.plotArea.groups;
		expect([bar?.element, scatter?.element]).toEqual(['barChart', 'scatterChart']);
		expect(scatter).toMatchObject({ scatterStyle: 'lineMarker', axisIds: bar!.axisIds });
		expect(saved.chartSpace.plotArea.axes).toEqual(source.plotArea.axes);
		const point = scatter!.series[0]!;
		const original = chartSeriesEntries(source)[1]!.series;
		expect(point.xValues).toEqual(original.categories);
		expect(point.yValues).toEqual(original.values);
		expect(point.categories).toBeUndefined();
		// Markers only, in the series colour, unsmoothed: Excel's XY (Scatter) series.
		expect(point.smooth).toBe(false);
		expect(point.spPr?.line?.fill).toEqual({ kind: 'none' });
		expect(point.marker).toMatchObject({
			symbol: 'circle',
			size: 5,
			spPr: { fill: original.spPr?.fill, line: { widthEmu: 9525, fill: original.spPr?.fill } },
		});
		const xml = writeChartSpace(changed);
		expect(xml).toMatch(/<c:scatterChart>.*<c:xVal>.*<c:yVal>.*<\/c:scatterChart>/su);
		expect(xml).not.toMatch(/<c:scatterChart>.*<c:cat>.*<\/c:scatterChart>/su);
		// Back to a column: the series takes its categories and values again.
		const back = setChartSeriesType(changed, 1, { kind: 'bar' });
		expect(back.plotArea.groups.map((group) => group.kind)).toEqual(['bar', 'bar']);
		expect(chartSeriesEntries(back)[1]!.series.categories).toEqual(original.categories);
		expect(chartSeriesEntries(back)[1]!.series.values).toEqual(original.values);
	});
	it('gives a combination scatter series its categories back on a whole-chart switch', () => {
		const source = columnChart();
		const original = chartSeriesEntries(source)[1]!.series;
		const combo = setChartSeriesType(source, 1, { kind: 'scatter' });
		const columns = setChartType(combo, { kind: 'bar', is3D: true });
		const series = chartSeriesEntries(columns)[1]!.series;
		expect(series.categories).toEqual(original.categories);
		expect(series.values).toEqual(original.values);
		expect(series.xValues).toBeUndefined();
	});
	it('keeps a scatter chart on two value axes a whole-chart type', () => {
		const scatter = setChartType(columnChart(), { kind: 'scatter' });
		expect(() => setChartSeriesType(scatter, 1, { kind: 'line' })).toThrow('axis pair');
	});
	it('joins equal groups anywhere in the plot area before writing, keeping other groups apart', () => {
		const source = columnChart();
		source.plotArea.groups[0]!.series.push(structuredClone(source.plotArea.groups[0]!.series[1]!));
		source.plotArea.groups[0]!.series[2]!.index = 2;
		source.plotArea.groups[0]!.series[2]!.order = 2;
		const split = setChartSeriesType(source, 1, { kind: 'line' });
		expect(split.plotArea.groups.map((group) => group.kind)).toEqual(['bar', 'line', 'bar']);
		const joined = joinChartGroups(split);
		expect(joined.plotArea.groups.map((group) => [group.kind, group.series.length])).toEqual([
			['bar', 2],
			['line', 1],
		]);
		expect(chartSeriesEntries(joined).map((entry) => entry.series.order)).toEqual([0, 2, 1]);
		const secondary = joinChartGroups(setChartSeriesSecondaryAxis(source, 1, true));
		expect(secondary.plotArea.groups.map((group) => group.series.length)).toEqual([2, 1]);
		expect(secondary.plotArea.groups[1]!.axisIds).not.toEqual(
			secondary.plotArea.groups[0]!.axisIds,
		);
		expect(joinChartGroups(source)).toEqual(source);
	});
	it('joins a neighbouring group of the new type, and only a neighbour', () => {
		const source = columnChart();
		source.plotArea.groups[0]!.series.push(structuredClone(source.plotArea.groups[0]!.series[1]!));
		source.plotArea.groups[0]!.series[2]!.index = 2;
		source.plotArea.groups[0]!.series[2]!.order = 2;
		const twoLines = setChartSeriesType(setChartSeriesType(source, 1, { kind: 'line' }), 2, {
			kind: 'line',
		});
		expect(twoLines.plotArea.groups.map((group) => [group.kind, group.series.length])).toEqual([
			['bar', 1],
			['line', 2],
		]);
		const between = setChartSeriesType(setChartSeriesType(source, 0, { kind: 'line' }), 2, {
			kind: 'line',
		});
		expect(between.plotArea.groups.map((group) => [group.kind, group.series.length])).toEqual([
			['line', 1],
			['bar', 1],
			['line', 1],
		]);
		expect(chartSeriesEntries(between).map((entry) => entry.series.index)).toEqual([0, 1, 2]);
	});
});
