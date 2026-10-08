import type { ChartPptxElement, PptxChartData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildHorizontalBarViewModel } from './chart-horizontal-bars';
import { buildChartViewModel } from './chart-view-model';
import type { ChartViewModel, SvgRect } from './chart-view-model';

const element = (chartData?: PptxChartData): ChartPptxElement =>
	({ id: 'c', type: 'chart', x: 0, y: 0, width: 400, height: 300, chartData }) as ChartPptxElement;

/** Columns A=[20, 10], B=[30, 40], C=[20, 20]: totals 70 and 70. */
const stacked = (overrides: Partial<PptxChartData> = {}): PptxChartData => ({
	chartType: 'bar',
	categories: ['Q1', 'Q2'],
	series: [
		{ name: 'A', values: [20, 10] },
		{ name: 'B', values: [30, 40] },
		{ name: 'C', values: [20, 20] },
	],
	grouping: 'stacked',
	style: { hasDataLabels: true },
	...overrides,
});

const axes = (min: number, max: number): PptxChartData['axes'] => [
	{ axisType: 'valAx', axPos: 'l', majorGridlines: true, min, max },
];

const rects = (vm: ChartViewModel): SvgRect[] =>
	vm.primitives.filter((p): p is SvgRect => p.kind === 'rect');

const labelTexts = (vm: ChartViewModel): string[] => vm.dataLabels.map((label) => label.text);

describe('stacked column data labels', () => {
	it('sits each label on its own stacked segment', () => {
		const vm = buildChartViewModel(element(stacked()));
		const bars = rects(vm);
		expect(vm.dataLabels).toHaveLength(bars.length);
		for (const [i, bar] of bars.entries()) {
			const label = vm.dataLabels[i];
			expect(label?.x).toBeCloseTo(bar.x + bar.w / 2, 6);
			expect(label?.y).toBeGreaterThanOrEqual(bar.y);
			expect(label?.y).toBeLessThanOrEqual(bar.y + bar.h);
		}
	});

	it('drops the label of a segment that c:max cuts away and centres a cut one', () => {
		// c:max 40: Q1's B (20..50) is cut to 20..40 and C (50..70) is gone;
		// Q2's B (10..50) is cut to 10..40 and C is gone.
		const vm = buildChartViewModel(element(stacked({ axes: axes(0, 40) })));
		expect(labelTexts(vm)).toStrictEqual(['20', '30', '10', '40']);
		const bars = rects(vm);
		const top = Math.min(...vm.gridlines.map((g) => g.y1));
		for (const [i, bar] of bars.entries()) {
			expect(bar.y).toBeGreaterThanOrEqual(top - 1e-6);
			expect(vm.dataLabels[i]?.y).toBeCloseTo(bar.y + bar.h / 2, 0);
		}
		// The cut segment's label is in the middle of what is left of it.
		expect(vm.dataLabels[1]?.y).toBeGreaterThan(top);
	});

	it('drops the label of a segment that c:min cuts away', () => {
		// c:min 25: both A segments (0..20 and 0..10) lie below the axis.
		const vm = buildChartViewModel(element(stacked({ axes: axes(25, 70) })));
		expect(labelTexts(vm)).toStrictEqual(['30', '20', '40', '20']);
	});
});

describe('stacked horizontal bar data labels', () => {
	it('drops a cut-away segment and centres a cut one', () => {
		const vm = buildHorizontalBarViewModel(
			element(),
			stacked({ barDirection: 'bar', axes: [{ axisType: 'valAx', axPos: 'b', min: 0, max: 40 }] }),
			['Q1', 'Q2'],
		);
		expect(labelTexts(vm)).toStrictEqual(['20', '30', '10', '40']);
		for (const [i, bar] of rects(vm).entries()) {
			expect(vm.dataLabels[i]?.x).toBeCloseTo(bar.x + bar.w / 2, 0);
		}
	});
});
