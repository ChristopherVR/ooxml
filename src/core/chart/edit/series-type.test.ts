import { describe, expect, it } from 'vitest';
import { columnChart } from './test-support/column-chart';
import { setChartSeriesType } from './series-type';
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
		expect(() => setChartSeriesType(source, 1, { kind: 'scatter' })).toThrow('vertical 2D');
		expect(() => setChartSeriesType(source, 1, { kind: 'line', is3D: true })).toThrow(
			'vertical 2D',
		);
		source.plotArea.groups[0]!.extLst = '<c:extLst/>';
		expect(() => setChartSeriesType(source, 1, { kind: 'line' })).toThrow('unknown group');
	});
});
