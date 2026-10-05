/**
 * A dotted combo line with hollow markers keeps its dash and marker outline
 * through Angular's vendored copy of `pptx-viewer-shared`. Like
 * `chart-renderer-legend-swatch.test.ts`, this checks the view model the
 * `pptx-chart-primitives` template binds, since the component is not
 * template-mounted in this package's tests.
 */
import type { ChartPptxElement, PptxChartData } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { buildChartViewModel } from './chart-renderer-helpers';

const comboData = {
	chartType: 'combo',
	categories: ['2023', '2024', '2025'],
	series: [
		{ name: 'Volume', values: [2.8, 1.8, 2.6], seriesChartType: 'bar', color: '#CDDCE5' },
		{
			name: 'Share',
			values: [3, 2, 3],
			seriesChartType: 'line',
			color: '#05507D',
			lineWidth: 1,
			lineDashStyle: 'sysDot',
			marker: {
				symbol: 'circle',
				size: 5,
				spPr: { fillColor: '#FFFFFF', strokeColor: '#05507D', strokeWidth: 1 },
			},
		},
	],
} as PptxChartData;

const element = {
	id: 'el-combo',
	type: 'chart',
	x: 0,
	y: 0,
	width: 400,
	height: 300,
	chartData: comboData,
} as ChartPptxElement;

describe('chart-renderer series line dash and marker outline (vendored shared)', () => {
	it('gives the line a dash array and the markers an outline', () => {
		const { primitives } = buildChartViewModel(element);
		const line = primitives.find((p) => p.kind === 'polyline');
		expect(line?.kind === 'polyline' && line.dashArray).toBeTruthy();
		const marker = primitives.find((p) => p.kind === 'circle');
		expect(marker).toMatchObject({ fill: '#FFFFFF', stroke: '#05507D' });
	});
});
