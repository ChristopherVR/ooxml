import { registerPptxWebControls } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { EditorState } from '../../editor/editor-state.svelte';
import ChartFormatPanel from './ChartFormatPanel.svelte';

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

function setup(selectChart = true) {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [CHART] }]);
	if (selectChart) {
		editor.select('chart1');
		editor.chartFormat.openPane('chart1', 'legend');
	}
	const target = document.createElement('div');
	document.body.append(target);
	const instance = mount(ChartFormatPanel, { target, props: { editor } });
	flushSync();
	cleanup = () => {
		void unmount(instance);
		target.remove();
	};
	return { editor, target };
}

const pane = (target: HTMLElement) => target.querySelector('pptx-ui-chart-format-pane');

describe('chart Format pane Svelte adapter', () => {
	it('shows the pane on the selected chart and target', () => {
		const { target } = setup();
		const node = pane(target)!;
		expect(node.target).toBe('legend');
		expect(node.disabled).toBe(false);
		expect(node.shadowRoot!.querySelector('h3')!.textContent).toBe('Format Legend');
	});

	it('applies a pane edit through the editor patch path and follows retargeting', () => {
		const { editor, target } = setup();
		const node = pane(target)!;
		const chartData = {
			...(CHART as Extract<PptxElement, { type: 'chart' }>).chartData!,
			style: { hasLegend: true, legendPosition: 'b' },
		};
		node.dispatchEvent(
			new CustomEvent('chart-format-change', { detail: { chartData, target: 'legend' } }),
		);
		const updated = editor.elementById('chart1') as Extract<PptxElement, { type: 'chart' }>;
		expect(updated.chartData?.style?.legendPosition).toBe('b');
		node.dispatchEvent(new CustomEvent('chart-format-target', { detail: { target: 'axis:1' } }));
		flushSync();
		expect(editor.chartFormat.targetFor('chart1')).toBe('axis:1');
		node.dispatchEvent(new CustomEvent('chart-format-close'));
		flushSync();
		expect(editor.chartFormat.open).toBe(false);
		expect(pane(target)).toBeNull();
	});

	it('renders nothing and closes when no chart is selected', () => {
		const { editor, target } = setup(false);
		editor.chartFormat.openPane('chart1');
		flushSync();
		expect(pane(target)).toBeNull();
		expect(editor.chartFormat.open).toBe(false);
	});
});
