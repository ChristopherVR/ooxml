// oxlint-disable react-hooks/rules-of-hooks
/**
 * Chart Format > Current Selection: the contextual tab draws the shared strip;
 * its drop-down retargets, Format Selection opens the shared Format pane, and a
 * committed pane edit or Reset to Match Style goes through the update path.
 */
import { mount } from '@vue/test-utils';
import type { ChartPptxElement, PptxChartData, PptxElement } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { defineComponent, h, nextTick, ref } from 'vue';

import { useChartFormatPane } from '../../composables/chart-format-pane';
import type { ChartFormatController } from '../../composables/chart-format-pane';
import ChartFormatPanel from '../inspector/ChartFormatPanel.vue';
import ContextualTabSection from './ContextualTabSection.vue';

function chart(): ChartPptxElement {
	return {
		id: 'chart-1',
		type: 'chart',
		x: 0,
		y: 0,
		width: 400,
		height: 300,
		chartData: {
			chartType: 'bar',
			categories: ['A', 'B'],
			series: [{ name: 'East', values: [1, 2], color: '#4472C4' }],
			style: { hasLegend: true },
			axes: [{ axisType: 'catAx' }, { axisType: 'valAx', majorGridlines: true }],
		},
	} as ChartPptxElement;
}

let unmount: (() => void) | null = null;
afterEach(() => {
	unmount?.();
	unmount = null;
	document.body.replaceChildren();
});

function mountStrip(canEdit = true) {
	const element = ref<PptxElement>(chart());
	const updateElement = vi.fn((_id: string, patch: Partial<PptxElement>) => {
		element.value = { ...element.value, ...patch } as PptxElement;
	});
	let controller!: ChartFormatController;
	const wrapper = mount(
		defineComponent({
			setup() {
				controller = useChartFormatPane({
					selectedElement: () => element.value,
					canEdit: () => canEdit,
					updateElement,
				});
				return () =>
					h('div', [h(ContextualTabSection, { tab: 'chartFormat' }), h(ChartFormatPanel)]);
			},
		}),
		{ attachTo: document.body },
	);
	unmount = () => wrapper.unmount();
	return { wrapper, controller, updateElement, element };
}

const pane = () => document.querySelector('pptx-ui-chart-format-pane');

describe('chart Format > Current Selection', () => {
	it('opens the Format pane on the current element from Format Selection', async () => {
		const { wrapper, controller } = mountStrip();
		expect(
			wrapper.find('[data-ribbon-group="chartFormat.currentSelection"]').exists(),
		).toBeTruthy();
		expect(pane()).toBeNull();
		controller.setTarget('legend');
		await nextTick();
		await wrapper.get('[data-testid="chart-format-selection"]').trigger('click');
		await nextTick();
		expect(controller.open.value).toBeTruthy();
		expect(pane()?.target).toBe('legend');
		expect(pane()?.shadowRoot?.querySelector('h3')?.textContent).toBe('Format Legend');
	});

	it('commits a pane edit and Reset to Match Style through the update path', async () => {
		const { wrapper, controller, updateElement } = mountStrip();
		controller.openPane('series:0');
		await nextTick();
		const next = {
			...chart().chartData,
			series: [{ name: 'East', values: [1, 2], color: '#FF0000' }],
		} as PptxChartData;
		pane()!.dispatchEvent(
			new CustomEvent('chart-format-change', { detail: { chartData: next, target: 'series:0' } }),
		);
		expect(updateElement).toHaveBeenCalledWith('chart-1', { chartData: next });
		await wrapper.get('[data-testid="chart-format-reset"]').trigger('click');
		const reset = updateElement.mock.calls.at(-1)?.[1] as { chartData: PptxChartData };
		expect(reset.chartData.series[0]?.color).toBeUndefined();
		expect(reset.chartData.elementFormats?.['series:0']).toStrictEqual({ reset: true });
		pane()!.dispatchEvent(new CustomEvent('chart-format-close'));
		await nextTick();
		expect(pane()).toBeNull();
	});

	it('disables Reset on a read-only deck', async () => {
		const { wrapper } = mountStrip(false);
		await nextTick();
		expect(wrapper.get('[data-testid="chart-format-reset"]').attributes('disabled')).toBeDefined();
	});
});
