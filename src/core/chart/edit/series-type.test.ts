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
		expect(() => setChartSeriesType(source, 1, { kind: 'bubble' })).toThrow('2D column');
		expect(() => setChartSeriesType(source, 1, { kind: 'pie' })).toThrow('2D column');
		expect(() => setChartSeriesType(source, 1, { kind: 'line', is3D: true })).toThrow('2D column');
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
		expect(back.plotArea.groups.map((group) => group.kind)).toEqual(['bar']);
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
	it('turns the x value axis of an XY chart into the shared category axis, as Excel saves it', () => {
		const scatter = setChartType(columnChart(), { kind: 'scatter' });
		const changed = setChartSeriesType(scatter, 1, { kind: 'line' });
		expect(changed.plotArea.groups.map((group) => group.kind)).toEqual(['scatter', 'line']);
		expect(changed.plotArea.groups.map((group) => group.axisIds)).toEqual([
			[11, 22],
			[11, 22],
		]);
		const [x, y] = changed.plotArea.axes;
		expect(x).toMatchObject({ kind: 'cat', id: 11, position: 'b' });
		expect(x!.crossBetween).toBeUndefined();
		expect(y).toMatchObject({ kind: 'val', crossBetween: 'between' });
		const line = chartSeriesEntries(changed)[1]!.series;
		expect(line.categories).toBeDefined();
		expect(line.xValues).toBeUndefined();
		expect(chartSeriesEntries(changed)[0]!.series.xValues).toBeDefined();
		expect(parseChartSpace(writeChartSpace(changed)).issues).toEqual([]);
	});
	it('moves a series that turns between horizontal bars and a vertical family to a secondary pair', () => {
		const bars = setChartType(columnChart(), { kind: 'bar', barDirection: 'bar' });
		const changed = setChartSeriesType(bars, 1, { kind: 'line' });
		expect(changed.plotArea.groups.map((group) => [group.kind, group.axisIds])).toEqual([
			['bar', [11, 22]],
			['line', [500000003, 500000004]],
		]);
		expect(changed.plotArea.axes.slice(2)).toMatchObject([
			{ kind: 'cat', id: 500000003, crossAxisId: 500000004, deleted: true, position: 'b' },
			{ kind: 'val', id: 500000004, crossAxisId: 500000003, position: 'r', crosses: 'max' },
		]);
		const column = setChartSeriesType(columnChart(), 1, { kind: 'bar', barDirection: 'bar' });
		expect(column.plotArea.groups.map((group) => group.barDirection)).toEqual(['col', 'bar']);
		expect(column.plotArea.axes.slice(2).map((axis) => axis.position)).toEqual(['l', 't']);
		// Back to bars: the series rejoins its group and the secondary pair goes.
		const back = setChartSeriesType(changed, 1, { kind: 'bar', barDirection: 'bar' });
		expect(back.plotArea.groups).toHaveLength(1);
		expect(back.plotArea.groups[0]!.series.map((series) => series.index)).toEqual([0, 1]);
		expect(back.plotArea.axes.map((axis) => axis.id)).toEqual([11, 22]);
		expect(() =>
			setChartSeriesType(setChartType(columnChart(), { kind: 'scatter' }), 1, {
				kind: 'bar',
				barDirection: 'bar',
			}),
		).toThrow(/horizontal bars/u);
	});
	it('rejoins the group a series left when its settings differ from a new group', () => {
		const source = columnChart();
		const round = setChartSeriesType(setChartSeriesType(source, 1, { kind: 'line' }), 1, {
			kind: 'bar',
		});
		expect(round.plotArea.groups).toHaveLength(1);
		expect(round.plotArea.groups[0]).toMatchObject({ gapWidth: 219, overlap: -27 });
		const { series: _round, ...joined } = round.plotArea.groups[0]!;
		const { series: _source, ...original } = source.plotArea.groups[0]!;
		expect(joined).toEqual(original);
		expect(round.plotArea.groups[0]!.series.map((series) => series.index)).toEqual([0, 1]);
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
