import type { ChartPptxElement, PptxChartData, PptxChartShapeProps } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildCartesianViewModel } from './chart-cartesian';
import { DEFAULT_CHART_TEXT_PX } from './chart-font';
import { buildHorizontalBarViewModel, valueToX } from './chart-horizontal-bars';
import type { SvgRect } from './chart-view-model';
import { estimateTextWidth } from './text-wrap-estimate';

function element(): ChartPptxElement {
	return {
		type: 'chart',
		id: 'c1',
		x: 0,
		y: 0,
		width: 480,
		height: 320,
	} as ChartPptxElement;
}

function chartData(overrides?: Partial<PptxChartData>): PptxChartData {
	return {
		chartType: 'bar',
		barDirection: 'bar',
		categories: ['A', 'B', 'C'],
		series: [{ name: 'S1', values: [4, 3, 5] }],
		style: { hasLegend: true },
		...overrides,
	};
}

const CATEGORIES = ['A', 'B', 'C'];

function rects(vm: { primitives: ReadonlyArray<{ kind: string }> }): SvgRect[] {
	return vm.primitives.filter((p): p is SvgRect => p.kind === 'rect');
}

describe('valueToX', () => {
	it('maps min to the left edge and max to the right edge', () => {
		const range = { min: 0, max: 10, span: 10 };
		expect(valueToX(0, range, 50, 450)).toBe(50);
		expect(valueToX(10, range, 50, 450)).toBe(450);
		expect(valueToX(5, range, 50, 450)).toBe(250);
	});

	it('honours reversed axis order', () => {
		const range = { min: 0, max: 10, span: 10, reverseOrder: true };
		expect(valueToX(0, range, 50, 450)).toBe(450);
		expect(valueToX(10, range, 50, 450)).toBe(50);
	});
});

describe('buildHorizontalBarViewModel', () => {
	it('draws bars that grow horizontally: width varies with value, height is constant', () => {
		const vm = buildHorizontalBarViewModel(element(), chartData(), ['A', 'B', 'C']);
		const bars = rects(vm);
		expect(bars).toHaveLength(3);
		const heights = new Set(bars.map((b) => b.h.toFixed(3)));
		expect(heights.size).toBe(1);
		// values 4, 3, 5: widths ordered accordingly.
		expect(bars[1].w).toBeLessThan(bars[0].w);
		expect(bars[0].w).toBeLessThan(bars[2].w);
	});

	it('emits vertical value gridlines and left-anchored category labels', () => {
		const vm = buildHorizontalBarViewModel(element(), chartData(), ['A', 'B', 'C']);
		for (const line of vm.gridlines) {
			expect(line.x1).toBe(line.x2);
			expect(line.y1).not.toBe(line.y2);
		}
		expect(vm.categoryLabels).toHaveLength(3);
		for (const label of vm.categoryLabels) {
			expect(label.textAnchor).toBe('end');
		}
		// Category labels sit left of the plot; value labels below it.
		const maxCatX = Math.max(...vm.categoryLabels.map((l) => l.x));
		const minGridX = Math.min(...vm.gridlines.map((l) => l.x1));
		expect(maxCatX).toBeLessThan(minGridX + 1);
	});

	it('draws value gridlines only when the parsed value axis has c:majorGridlines', () => {
		const withAxes = (majorGridlines: boolean) =>
			buildHorizontalBarViewModel(
				element(),
				chartData({
					axes: [
						{ axisType: 'catAx', axPos: 'l', axisId: 1, crossAxisId: 2 },
						{ axisType: 'valAx', axPos: 'b', axisId: 2, crossAxisId: 1, majorGridlines },
					],
				}),
				CATEGORIES,
			);
		const hidden = withAxes(false);
		const shown = withAxes(true);
		expect(hidden.gridlines).toStrictEqual([]);
		expect(shown.gridlines.length).toBeGreaterThan(0);
		expect(hidden.axisLabels).toStrictEqual(shown.axisLabels);
	});

	it('takes the value gridline style from c:majorGridlines/c:spPr', () => {
		const withGridlineLine = (majorGridlinesSpPr: PptxChartShapeProps) =>
			buildHorizontalBarViewModel(
				element(),
				chartData({
					axes: [{ axisType: 'valAx', axPos: 'b', majorGridlines: true, majorGridlinesSpPr }],
				}),
				CATEGORIES,
			);
		const styled = withGridlineLine({ strokeColor: '#D9D9D9', strokeWidth: 0.75 });
		expect(styled.gridlines.length).toBeGreaterThan(0);
		expect(styled.gridlines.every((line) => line.stroke === '#D9D9D9')).toBeTruthy();
		expect(withGridlineLine({ lineNoFill: true }).gridlines).toStrictEqual([]);
	});

	it('stacks series along x when grouping is stacked', () => {
		const vm = buildHorizontalBarViewModel(
			element(),
			chartData({
				grouping: 'stacked',
				series: [
					{ name: 'S1', values: [2, 2] },
					{ name: 'S2', values: [3, 1] },
				],
			}),
			['A', 'B'],
		);
		const bars = rects(vm);
		expect(bars).toHaveLength(4);
		const catABars = bars.filter((b) => b.part?.pointIndex === 0);
		expect(catABars).toHaveLength(2);
		// Second segment starts where the first ends.
		const [first, second] = catABars;
		expect(second.x).toBeCloseTo(first.x + first.w, 1);
	});

	it('normalises percentStacked categories to the full plot width', () => {
		const vm = buildHorizontalBarViewModel(
			element(),
			chartData({
				grouping: 'percentStacked',
				series: [
					{ name: 'S1', values: [1, 9] },
					{ name: 'S2', values: [3, 1] },
				],
			}),
			['A', 'B'],
		);
		const bars = rects(vm);
		const catWidth = (pointIndex: number) =>
			bars.filter((b) => b.part?.pointIndex === pointIndex).reduce((sum, b) => sum + b.w, 0);
		expect(catWidth(0)).toBeCloseTo(catWidth(1), 0);
	});

	it('sizes stacked and percentStacked bars from c:gapWidth like a one-bar cluster', () => {
		const stacked = (grouping: 'stacked' | 'percentStacked', barGapWidth?: number) => {
			const vm = buildHorizontalBarViewModel(
				element(),
				chartData({
					grouping,
					...(barGapWidth !== undefined ? { barGapWidth } : {}),
					series: [
						{ name: 'S1', values: [2, 2] },
						{ name: 'S2', values: [3, 1] },
					],
				}),
				['A', 'B'],
			);
			const band = (vm.gridlines[0].y2 - vm.gridlines[0].y1) / 2;
			return { heights: rects(vm).map((bar) => bar.h), band };
		};
		for (const grouping of ['stacked', 'percentStacked'] as const) {
			const sized = stacked(grouping, 50);
			for (const h of sized.heights) {
				expect(h).toBeCloseTo(sized.band / 1.5, 6);
			}
		}
		// Without c:gapWidth both take the ECMA-376 default of 150%: 40% of the band.
		for (const grouping of ['stacked', 'percentStacked'] as const) {
			const fallback = stacked(grouping);
			for (const h of fallback.heights) {
				expect(h).toBeCloseTo(fallback.band / 2.5, 6);
			}
		}
	});

	it('cuts stacked bars at the plot edges for c:min above 0 and c:max below a total', () => {
		const vm = buildHorizontalBarViewModel(
			element(),
			chartData({
				grouping: 'stacked',
				axes: [{ axisType: 'valAx', axPos: 'b', majorGridlines: true, min: 3, max: 4 }],
				series: [
					{ name: 'S1', values: [2, 2] },
					{ name: 'S2', values: [3, 1] },
				],
			}),
			['A', 'B'],
		);
		const left = Math.min(...vm.gridlines.map((line) => line.x1));
		const right = Math.max(...vm.gridlines.map((line) => line.x1));
		const bars = rects(vm);
		// S1 (0..2) lies below 3 in both categories and is dropped. S2 runs 2..5
		// in A, cut to the whole 3..4 axis, and 2..3 in B, which touches only the edge.
		expect(bars).toHaveLength(1);
		expect(bars[0].x).toBeCloseTo(left, 6);
		expect(bars[0].x + bars[0].w).toBeCloseTo(right, 6);
	});

	it('honours explicit c:min and c:max on a stacked bar chart value axis', () => {
		const vm = buildHorizontalBarViewModel(
			element(),
			chartData({
				grouping: 'stacked',
				axes: [{ axisType: 'valAx', axPos: 'b', majorGridlines: true, min: 0, max: 20 }],
				series: [
					{ name: 'S1', values: [2, 2] },
					{ name: 'S2', values: [3, 1] },
				],
			}),
			['A', 'B'],
		);
		const left = Math.min(...vm.gridlines.map((line) => line.x1));
		const right = Math.max(...vm.gridlines.map((line) => line.x1));
		const totalA = rects(vm)
			.filter((bar) => bar.part?.pointIndex === 0)
			.reduce((sum, bar) => sum + bar.w, 0);
		// Category A sums to 5 of an axis that runs to 20.
		expect(totalA / (right - left)).toBeCloseTo(0.25, 2);
		expect(vm.axisLabels.some((label) => label.text === '20')).toBeTruthy();
	});

	it.each(['stacked', 'percentStacked'] as const)(
		'reverses a %s bar chart value axis for c:orientation maxMin',
		(grouping) => {
			const vm = buildHorizontalBarViewModel(
				element(),
				chartData({
					grouping,
					axes: [{ axisType: 'valAx', axPos: 'b', majorGridlines: true, orientation: 'maxMin' }],
					series: [
						{ name: 'S1', values: [2, 2] },
						{ name: 'S2', values: [3, 1] },
					],
				}),
				['A', 'B'],
			);
			const [first, second] = rects(vm).filter((bar) => bar.part?.pointIndex === 0);
			// Values grow leftwards: the second segment sits left of the first.
			expect(second.x + second.w).toBeCloseTo(first.x, 1);
		},
	);

	it.each([
		['mixed-sign', [30, -10], [-70, 90]],
		['all-negative', [-30, -10], [-70, -90]],
	] as const)('keeps every %s percentStacked segment on the plot', (_name, a, b) => {
		const vm = buildHorizontalBarViewModel(
			element(),
			chartData({
				grouping: 'percentStacked',
				style: { hasDataLabels: true },
				series: [
					{ name: 'S1', values: [...a] },
					{ name: 'S2', values: [...b] },
				],
			}),
			['A', 'B'],
		);
		const left = Math.min(...vm.gridlines.map((line) => line.x1));
		const right = Math.max(...vm.gridlines.map((line) => line.x1));
		const bars = rects(vm);
		expect(bars).toHaveLength(4);
		for (const bar of bars) {
			expect(bar.x).toBeGreaterThanOrEqual(left - 1e-6);
			expect(bar.x + bar.w).toBeLessThanOrEqual(right + 1e-6);
		}
		expect(vm.dataLabels).toHaveLength(4);
		expect(vm.zeroLine).toBeDefined();
	});

	it('draws a vertical zero line when the range spans zero', () => {
		const vm = buildHorizontalBarViewModel(
			element(),
			chartData({ series: [{ name: 'S1', values: [-2, 3] }] }),
			['A', 'B'],
		);
		expect(vm.zeroLine).toBeDefined();
		expect(vm.zeroLine?.x1).toBe(vm.zeroLine?.x2);
	});

	it('reserves enough left margin that a category label is not clipped against the SVG edge', () => {
		// Regression for the clipping a freshly inserted "Insert > Chart > Bar"
		// showed: default category text ("Category 1"/2/3) is far wider than the
		// 40px band sized for a numeric value axis, and rendered as "egory 1"
		// because the label's own text ran past the SVG's left edge (x=0).
		const categoryLabels = ['Category 1', 'Category 2', 'Category 3'];
		const vm = buildHorizontalBarViewModel(
			element(),
			chartData({ categories: categoryLabels }),
			categoryLabels,
		);
		expect(vm.categoryLabels).toHaveLength(3);
		for (const label of vm.categoryLabels) {
			expect(label.textAnchor).toBe('end');
			const estimatedWidth = estimateTextWidth(label.text, DEFAULT_CHART_TEXT_PX);
			// End-anchored text at `label.x` draws leftward; its left-most ink must
			// stay at or right of the SVG's own left edge (x=0).
			expect(label.x - estimatedWidth).toBeGreaterThanOrEqual(0);
		}
	});

	it('widens the left plot inset for longer category labels than for short ones', () => {
		const short = buildHorizontalBarViewModel(element(), chartData(), ['A', 'B', 'C']);
		const longLabels = ['Category 1', 'Category 2', 'Category 3'];
		const long = buildHorizontalBarViewModel(
			element(),
			chartData({ categories: longLabels }),
			longLabels,
		);
		const plotLeft = (vm: { gridlines: ReadonlyArray<{ x1: number }> }) =>
			Math.min(...vm.gridlines.map((g) => g.x1));
		expect(plotLeft(long)).toBeGreaterThan(plotLeft(short));
	});

	it('matches the COM-verified bar-to-band ratio for a 3-series clustered chart (gapWidth=219, overlap=-27)', () => {
		// Same COM-verified fixture as chart-cartesian-bars.test.ts (Office's own
		// default 3-series clustered column chart), mirrored onto the horizontal
		// (bar) direction: the bar is 17.6% of the category band.
		const vm = buildHorizontalBarViewModel(
			element(),
			chartData({
				barGapWidth: 219,
				barOverlap: -27,
				series: [
					{ name: 'S1', values: [4.3, 2.5, 3.5, 4.5] },
					{ name: 'S2', values: [2.4, 4.4, 1.8, 2.8] },
					{ name: 'S3', values: [2, 2, 3, 5] },
				],
			}),
			['A', 'B', 'C', 'D'],
		);
		const bars = rects(vm);
		expect(bars).toHaveLength(12);
		// The band per category is the y-distance between the first series' bar
		// in consecutive categories (pointIndex 0 vs. pointIndex 1).
		const firstOfCategory = (pointIndex: number) =>
			bars.find((b) => b.part?.seriesIndex === 0 && b.part?.pointIndex === pointIndex);
		const band = (firstOfCategory(1)?.y ?? 0) - (firstOfCategory(0)?.y ?? 0);
		expect(bars[0].h / band).toBeCloseTo(0.176, 2);
	});

	it('is dispatched by the cartesian builder for barDirection "bar" only', () => {
		const horizontal = buildCartesianViewModel(element(), chartData(), ['A', 'B', 'C'], 'bar');
		const vertical = buildCartesianViewModel(
			element(),
			chartData({ barDirection: 'col' }),
			['A', 'B', 'C'],
			'bar',
		);
		// The horizontal build carries end-anchored side category labels; the
		// column build centres its labels under the plot.
		expect(horizontal.categoryLabels.every((l) => l.textAnchor === 'end')).toBeTruthy();
		expect(vertical.categoryLabels.every((l) => l.textAnchor === 'middle')).toBeTruthy();
		const verticalBars = rects(vertical);
		const heights = new Set(verticalBars.map((b) => b.h.toFixed(3)));
		expect(heights.size).toBeGreaterThan(1);
	});
});
