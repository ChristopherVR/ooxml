import { registerPptxWebControls } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { EditorState } from '../../../editor/editor-state.svelte';
import ChartSelectionStrip from './ChartSelectionStrip.svelte';

registerPptxWebControls();
const CHART = {
	type: 'chart',
	id: 'chart1',
	x: 0,
	y: 0,
	width: 480,
	height: 300,
	chartData: {
		chartType: 'bar',
		categories: ['Q1', 'Q2'],
		series: [{ name: 'East', values: [1, 2], color: '#4472C4' }],
		style: { hasLegend: true },
		axes: [
			{ axisType: 'catAx', axisId: 1 },
			{ axisType: 'valAx', axisId: 2, majorGridlines: true },
		],
	},
} as PptxElement;
let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function setup() {
	const chart = {
		...CHART,
		chartData: {
			...(CHART as Extract<PptxElement, { type: 'chart' }>).chartData!,
			series: [{ name: 'East', values: [1, 2], color: '#C00000' }],
		},
	} as PptxElement;
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [chart] }]);
	editor.select('chart1');
	const target = document.createElement('div');
	document.body.append(target);
	const instance = mount(ChartSelectionStrip, { target, props: { editor } });
	flushSync();
	cleanup = () => {
		void unmount(instance);
		target.remove();
	};
	const strip = target.querySelector('pptx-ui-ribbon-home-chart-selection')!;
	const request = (id: string, value?: string) =>
		strip.dispatchEvent(
			new CustomEvent('home-request', { detail: value === undefined ? { id } : { id, value } }),
		);
	return { editor, request };
}

describe('chart Format > Current Selection (Svelte)', () => {
	it('selects an element, opens the pane and resets the element to match the style', () => {
		const { editor, request } = setup();
		request('chartFormat.currentSelection.chartElements', 'series:0');
		expect(editor.chartFormat.targetFor('chart1')).toBe('series:0');
		request('chartFormat.currentSelection.formatSelection');
		expect(editor.chartFormat.open).toBe(true);
		request('chartFormat.currentSelection.resetToMatchStyle');
		const chart = editor.elementById('chart1') as Extract<PptxElement, { type: 'chart' }>;
		expect(chart.chartData?.series[0]?.color).toBeUndefined();
		expect(chart.chartData?.elementFormats?.['series:0']).toStrictEqual({ reset: true });
	});
});
