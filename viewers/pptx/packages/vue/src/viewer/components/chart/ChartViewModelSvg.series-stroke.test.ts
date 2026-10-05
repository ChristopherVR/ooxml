import { mount } from '@vue/test-utils';
import type { ChartPptxElement, PptxChartData } from 'pptx-viewer-core';
import { buildChartViewModel } from 'pptx-viewer-shared';
import { describe, expect, it } from 'vitest';

import ChartViewModelSvg from './ChartViewModelSvg.vue';

/** A dotted combo line with hollow markers keeps its dash and marker outline. */
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

describe('chartViewModelSvg: series line dash and marker outline', () => {
	it('writes stroke-dasharray on the line and an outline on each marker', () => {
		const wrapper = mount(ChartViewModelSvg, {
			props: { elementId: 'c1', vm: buildChartViewModel(element) },
		});
		expect(wrapper.find('polyline').attributes('stroke-dasharray')).toBeTruthy();
		const marker = wrapper.find('circle');
		expect(marker.attributes('fill')).toBe('#FFFFFF');
		expect(marker.attributes('stroke')).toBe('#05507D');
	});
});
