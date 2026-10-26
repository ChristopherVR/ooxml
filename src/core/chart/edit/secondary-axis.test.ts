import { describe, expect, it } from 'vitest';
import { chartSeriesEntries } from './draft';
import { chartSeriesSecondaryAxisSupported, setChartSeriesSecondaryAxis } from './secondary-axis';
import { columnChart } from './test-support/column-chart';
import { writeChartSpace } from '../write-space';
import { parseChartSpace } from '../parse-space';
import { renderChartSpaceSvg } from '../render/render-space';

const reopen = (space: ReturnType<typeof columnChart>) =>
	parseChartSpace(writeChartSpace(space)).chartSpace;
const entries = (space: ReturnType<typeof columnChart>) =>
	chartSeriesEntries(space).map(({ series }) => series);
const svg = (space: ReturnType<typeof columnChart>) =>
	renderChartSpaceSvg(space, { width: 480, height: 300 });

describe('secondary axes with stable document-order selections', () => {
	it('splits a normal column group without changing authored ids, data references or target positions', () => {
		const source = columnChart();
		const before = structuredClone(source);
		const group = source.plotArea.groups[0]!;
		const third = structuredClone(group.series[0]!);
		third.index = 7;
		third.order = 9;
		group.series.push(third);
		const moved = reopen(setChartSeriesSecondaryAxis(source, 1, true));
		expect(entries(moved)).toEqual(entries(reopen(source)));
		expect(moved.plotArea.groups.map((g) => g.series.length)).toEqual([1, 1, 1]);
		expect(moved.plotArea.groups.every((g) => g.gapWidth === 219 && g.overlap === -27)).toBe(true);
		for (const index of [0, 1, 2])
			expect(
				svg(moved).match(new RegExp(`data-chart-series="${index}" data-chart-point`, 'g')),
			).toHaveLength(3);
		expect(before.plotArea.groups[0]!.series).toHaveLength(2);
		expect(source.plotArea.axes).toHaveLength(2);
	});
	it('restores the original clustered bar geometry on return to the primary axis', () => {
		const source = columnChart();
		const restored = setChartSeriesSecondaryAxis(
			setChartSeriesSecondaryAxis(source, 1, true),
			1,
			false,
		);
		const rectangles = (s: string) =>
			[...s.matchAll(/<g data-chart-series="[01]" data-chart-point="\d"><rect [^>]*>/g)].map(
				(m) => m[0],
			);
		expect(rectangles(svg(restored))).toEqual(rectangles(svg(source)));
	});
	it('refuses the primary anchor and chart families the combination painter cannot represent', () => {
		expect(chartSeriesSecondaryAxisSupported(columnChart(), 0)).toBe(false);
		for (const edit of [
			{ barDirection: 'bar' },
			{ grouping: 'stacked' },
			{ extLst: '<c:extLst/>' },
		] as const) {
			const source = columnChart();
			Object.assign(source.plotArea.groups[0]!, edit);
			expect(chartSeriesSecondaryAxisSupported(source, 1)).toBe(false);
			expect(() => setChartSeriesSecondaryAxis(source, 1, true)).toThrow();
		}
	});
});
