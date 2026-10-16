import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';

import { chartFormatDoubleClick, chartFormatPointerDown } from '../components/chart-format-pointer';
import { ChartFormatState } from './chart-format-state.svelte';
import {
	buildEditorContextMenuEntries,
	recordChartContextTarget,
	runContextMenuCommand,
} from './context-menu-dispatch';
import { EditorState } from './editor-state.svelte';

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
		series: [
			{ name: 'East', values: [1, 2], color: '#4472C4' },
			{ name: 'West', values: [3, 4], color: '#ED7D31' },
		],
		style: { hasLegend: true },
		axes: [
			{ axisType: 'catAx', axisId: 1 },
			{ axisType: 'valAx', axisId: 2, majorGridlines: true },
		],
	},
} as PptxElement;

function makeEditor(): EditorState {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [CHART] }]);
	editor.select('chart1');
	return editor;
}

afterEach(() => document.body.replaceChildren());

describe('chart format state', () => {
	it('keeps a target per chart and reads the chart area for any other chart', () => {
		const state = new ChartFormatState();
		expect(state.targetFor('a')).toBe('chartArea');
		state.setTarget('a', 'series:1');
		expect(state.targetFor('a')).toBe('series:1');
		expect(state.targetFor('b')).toBe('chartArea');
		state.openPane('b');
		expect(state.open).toBe(true);
		expect(state.targetFor('b')).toBe('chartArea');
		state.openPane('b', 'legend');
		expect(state.targetFor('b')).toBe('legend');
		state.close();
		expect(state.open).toBe(false);
	});
});

describe('chart context menu', () => {
	it('names the entry after the right-clicked element and opens the pane on it', () => {
		const editor = makeEditor();
		const host = document.createElement('div');
		host.dataset.elementId = 'chart1';
		host.innerHTML =
			'<svg viewBox="0 0 480 300"><rect data-chart-part="dataPoint" data-chart-series="1" data-chart-point="0"></rect></svg>';
		document.body.append(host);
		recordChartContextTarget(editor, 'chart1', {
			target: host.querySelector('rect'),
			clientX: 0,
			clientY: 0,
		});
		expect(editor.chartFormat.targetFor('chart1')).toBe('series:1');
		const entry = buildEditorContextMenuEntries({ editor }).find(
			(e) => e.id === 'format-chart-element',
		);
		expect(entry?.labelKey).toBe('pptx.contextMenu.formatDataSeries');
		runContextMenuCommand('format-chart-element', { editor });
		expect(editor.chartFormat.open).toBe(true);
		expect(editor.chartFormat.targetFor('chart1')).toBe('series:1');
	});

	it('ignores a right-click on anything but a chart', () => {
		const editor = makeEditor();
		recordChartContextTarget(editor, 'missing', { target: null, clientX: 0, clientY: 0 });
		expect(editor.chartFormat.targetFor('missing')).toBe('chartArea');
	});
});

describe('chart canvas pointer', () => {
	it('targets on press and opens the pane on double-click, leaving the title alone', () => {
		const state = new ChartFormatState();
		const root = document.createElement('div');
		root.innerHTML =
			'<svg viewBox="0 0 480 300"><rect data-chart-part="series" data-chart-series="0"></rect><text data-chart-part="title">T</text></svg>';
		document.body.append(root);
		const mark = root.querySelector('rect')!;
		const title = root.querySelector('text')!;
		const event = (target: Element) => ({
			target,
			clientX: 0,
			clientY: 0,
			stopped: false,
			stopPropagation() {
				this.stopped = true;
			},
		});
		chartFormatPointerDown(event(mark), root, CHART, state);
		expect(state.targetFor('chart1')).toBe('series:0');
		const onTitle = event(title);
		expect(chartFormatDoubleClick(onTitle, root, CHART, state)).toBe(false);
		expect(onTitle.stopped).toBe(false);
		expect(state.open).toBe(false);
		const onMark = event(mark);
		expect(chartFormatDoubleClick(onMark, root, CHART, state)).toBe(true);
		expect(onMark.stopped).toBe(true);
		expect(state.open).toBe(true);
	});
});
