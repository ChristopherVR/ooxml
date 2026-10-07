import type { ChartPptxElement, PptxChartData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { CHART_PX_PER_PT } from './chart-font';
import type { SvgPolyline } from './chart-view-model';
import { buildChartViewModel } from './chart-view-model';
import { buildDashArray } from './connector-dash';

function lineChart(series: PptxChartData['series']): ChartPptxElement {
	return {
		id: 'el-line',
		type: 'chart',
		x: 0,
		y: 0,
		width: 400,
		height: 300,
		chartData: { chartType: 'line', categories: ['A', 'B', 'C'], series },
	} as ChartPptxElement;
}

function polylines(element: ChartPptxElement): SvgPolyline[] {
	return buildChartViewModel(element).primitives.filter(
		(p): p is SvgPolyline => p.kind === 'polyline',
	);
}

describe('line chart series stroke', () => {
	it('draws the series a:ln width and dash', () => {
		const [line] = polylines(
			lineChart([{ name: 'S', values: [1, 2, 3], lineWidth: 1.5, lineDashStyle: 'dash' }]),
		);
		const px = 1.5 * CHART_PX_PER_PT;
		expect(line).toMatchObject({ strokeWidth: px, dashArray: buildDashArray('dash', px) });
	});

	it('keeps the 2.4px solid default for a series without a:ln', () => {
		const [line] = polylines(lineChart([{ name: 'S', values: [1, 2, 3] }]));
		expect(line.strokeWidth).toBe(2.4);
		expect(line.dashArray).toBeUndefined();
	});
});
