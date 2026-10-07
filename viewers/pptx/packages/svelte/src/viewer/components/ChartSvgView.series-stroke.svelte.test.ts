import { buildChartViewModel } from 'ooxml-ui/pptx';
import type { ChartPptxElement, PptxChartData } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import ChartSvgView from './ChartSvgView.svelte';

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

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

describe('chartSvgView series line dash and marker outline', () => {
	it('writes stroke-dasharray on the line and an outline on each marker', () => {
		const target = document.createElement('div');
		document.body.appendChild(target);
		const component = mount(ChartSvgView, {
			target,
			props: { vm: buildChartViewModel(element), preserveAspectRatio: 'none', legendItems: [] },
		});
		flushSync();
		cleanup = () => {
			void unmount(component);
			target.remove();
		};
		expect(target.querySelector('polyline')?.getAttribute('stroke-dasharray')).toBeTruthy();
		const marker = target.querySelector('circle');
		expect(marker?.getAttribute('fill')).toBe('#FFFFFF');
		expect(marker?.getAttribute('stroke')).toBe('#05507D');
	});
});
