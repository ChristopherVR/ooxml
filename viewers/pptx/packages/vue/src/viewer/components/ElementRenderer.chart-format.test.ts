import { mount } from '@vue/test-utils';
import type { ChartPptxElement, PptxElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';
import { ref } from 'vue';

import type { ChartFormatController } from '../composables/chart-format-pane';
import { ChartFormatKey } from '../composables/chart-format-pane';
import type { ChartCanvasEditContext } from '../composables/chart-part-selection';
import { ChartCanvasEditKey } from '../composables/chart-part-selection';
import ElementRenderer from './ElementRenderer.vue';

/**
 * Chart Format on the canvas: pressing a chart element makes it the current
 * selection and double-clicking it opens its Format pane, while a double-click
 * on the title still edits the title in place.
 */
const CHART = {
	id: 'ch_1',
	type: 'chart',
	x: 0,
	y: 0,
	width: 400,
	height: 300,
	chartData: {
		chartType: 'bar',
		categories: ['Q1', 'Q2'],
		series: [
			{ name: 'Revenue', values: [100, 150] },
			{ name: 'Cost', values: [80, 90] },
		],
		title: 'Sales',
		style: { hasTitle: true },
	},
} as ChartPptxElement;

function mountChart() {
	const setTarget = vi.fn();
	const openPane = vi.fn();
	const ctx: ChartCanvasEditContext = {
		selection: ref(null),
		setSelection: () => {},
		canSelectCharts: () => true,
		canEditChart: () => true,
		updateElement: () => {},
	};
	const chartFormat = { setTarget, openPane } as unknown as ChartFormatController;
	const wrapper = mount(ElementRenderer, {
		props: {
			element: CHART as PptxElement,
			mediaDataUrls: new Map<string, string>(),
			zIndex: 1,
			interactive: true,
		},
		global: {
			provide: {
				[ChartCanvasEditKey as symbol]: ctx,
				[ChartFormatKey as symbol]: chartFormat,
			},
		},
	});
	return { wrapper, setTarget, openPane };
}

describe('elementRenderer - chart Format targeting', () => {
	it('targets the pressed series and opens its Format pane on double-click', () => {
		const { wrapper, setTarget, openPane } = mountChart();
		const bar = wrapper.element.querySelector(
			"rect[data-chart-part='dataPoint'][data-chart-series='1']",
		)!;
		bar.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, cancelable: true }));
		expect(setTarget).toHaveBeenCalledWith('series:1');
		bar.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		expect(openPane).toHaveBeenCalledWith('series:1');
	});

	it('keeps the inline title editor on a title double-click', async () => {
		const { wrapper, openPane } = mountChart();
		const title = wrapper.element.querySelector("[data-chart-part='title']")!;
		title.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		await wrapper.vm.$nextTick();
		expect(openPane).not.toHaveBeenCalled();
		expect(wrapper.element.querySelector('input')).not.toBeNull();
	});
});
