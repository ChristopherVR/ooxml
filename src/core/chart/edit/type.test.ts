import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { setChartGrouping, setChartType } from './type';
import { columnChart } from './test-support/column-chart';

/** Written and read back with no issue, so the switched part is schema-ordered and complete. */
const roundTrip = (space: ReturnType<typeof columnChart>) => {
	const result = parseChartSpace(writeChartSpace(space));
	expect(result.issues).toEqual([]);
	return result.chartSpace;
};

describe('setChartType', () => {
	it('switches a column chart to a line chart, keeping series and axes', () => {
		const before = columnChart();
		const line = setChartType(before, { kind: 'line' });
		expect(before.plotArea.groups[0]!.kind).toBe('bar');
		const group = line.plotArea.groups[0]!;
		expect(group).toMatchObject({
			kind: 'line',
			element: 'lineChart',
			grouping: 'standard',
			marker: true,
			axisIds: [11, 22],
		});
		expect(group.gapWidth).toBeUndefined();
		expect(group.overlap).toBeUndefined();
		expect(group.series.map((s) => s.tx?.text)).toEqual(['Plan', 'Actual']);
		expect(group.series[0]!.invertIfNegative).toBeUndefined();
		// The bar colour becomes the line colour, with Excel's series line width and cap.
		expect(group.series[0]!.spPr?.line).toMatchObject({
			widthEmu: 28575,
			cap: 'rnd',
			fill: { kind: 'solid', color: { kind: 'srgb', value: '4472C4', transforms: [] } },
		});
		expect(line.plotArea.axes[1]!.crossBetween).toBe('between');
		expect(line.plotArea.axes.map((a) => a.id)).toEqual([11, 22]);
		expect(roundTrip(line).plotArea.groups[0]!.element).toBe('lineChart');
	});

	it('drops the axes when switching to a pie and adds Excel default axes back', () => {
		const pie = setChartType(columnChart(), { kind: 'pie' });
		expect(pie.plotArea.axes).toEqual([]);
		expect(pie.plotArea.groups[0]).toMatchObject({
			element: 'pieChart',
			varyColors: true,
			firstSliceAngle: 0,
			axisIds: [],
		});
		roundTrip(pie);
		const column = setChartType(pie, { kind: 'bar', barDirection: 'col' });
		expect(column.plotArea.axes.map((a) => [a.kind, a.id, a.position])).toEqual([
			['cat', 500000001, 'b'],
			['val', 500000002, 'l'],
		]);
		expect(column.plotArea.groups[0]).toMatchObject({ grouping: 'clustered', gapWidth: 150 });
		roundTrip(column);
	});

	it('maps categories and values to x and y values for a scatter chart and back', () => {
		const scatter = setChartType(columnChart(), { kind: 'scatter' });
		const series = scatter.plotArea.groups[0]!.series[0]!;
		expect(series.categories).toBeUndefined();
		expect(series.xValues?.formula).toBe('Sheet1!$A$2:$A$4');
		expect(series.yValues?.formula).toBe('Sheet1!$B$2:$B$4');
		expect(scatter.plotArea.axes.map((a) => a.kind)).toEqual(['val', 'val']);
		expect(scatter.plotArea.axes[0]!.auto).toBeUndefined();
		expect(scatter.plotArea.axes.map((a) => a.crossBetween)).toEqual(['midCat', 'midCat']);
		roundTrip(scatter);
		const back = setChartType(scatter, { kind: 'bar' });
		expect(back.plotArea.axes.map((a) => a.kind)).toEqual(['cat', 'val']);
		expect(back.plotArea.groups[0]!.series[0]!.values?.formula).toBe('Sheet1!$B$2:$B$4');
		roundTrip(back);
	});

	it('rotates the axes when columns become horizontal bars', () => {
		const bars = setChartType(columnChart(), { kind: 'bar', barDirection: 'bar' });
		expect(bars.plotArea.axes.map((a) => a.position)).toEqual(['l', 'b']);
		expect(bars.plotArea.groups[0]).toMatchObject({ gapWidth: 219, overlap: -27 });
	});

	it('demotes clustered grouping and overlaps stacked bars', () => {
		const area = setChartType(columnChart(), { kind: 'area', grouping: 'clustered' });
		expect(area.plotArea.groups[0]!.grouping).toBe('standard');
		const stacked = setChartGrouping(columnChart(), 'stacked');
		expect(stacked.plotArea.groups[0]).toMatchObject({ grouping: 'stacked', overlap: 100 });
	});

	it('adds a 3-D view for a 3-D chart and removes it again', () => {
		const bar3D = setChartType(columnChart(), { kind: 'bar', is3D: true });
		expect(bar3D.plotArea.groups[0]!.element).toBe('bar3DChart');
		expect(bar3D.view3D).toEqual({ rotX: 15, rotY: 20, rightAngleAxes: true });
		roundTrip(bar3D);
		expect(setChartType(bar3D, { kind: 'line' }).view3D).toBeUndefined();
	});

	it('rejects families it cannot build', () => {
		expect(() => setChartType(columnChart(), { kind: 'stock' })).toThrow(RangeError);
		const line = setChartType(columnChart(), { kind: 'line', is3D: true });
		expect(line.plotArea.axes.map((axis) => axis.kind)).toEqual(['cat', 'val', 'ser']);
		expect(roundTrip(line).plotArea.groups[0]!.element).toBe('line3DChart');
		const surface = setChartType(setChartType(columnChart(), { kind: 'pie' }), {
			kind: 'surface',
			is3D: true,
		});
		expect(surface.plotArea.axes.map((axis) => axis.kind)).toEqual(['cat', 'val', 'ser']);
		expect(new Set(surface.plotArea.axes.map((axis) => axis.id)).size).toBe(3);
		roundTrip(surface);
	});
});
