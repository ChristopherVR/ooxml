// @vitest-environment happy-dom
import type { ChartPptxElement, PptxElement } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';

import { createTranslator } from '../i18n';
import { CHART_FORMAT_REQUEST_EVENT } from '../render/elements/chart-editable';
import { createInitialViewerState, createStore } from '../state';
import type { ViewerState } from '../state';
import { createChartFormatPaneHost } from './chart-format-pane-host';
import type { ApplyToSelected } from './editor-apply-to-selected';
import { chartFormatTargetOf, createChartFormatActions } from './editor-chart-format-actions';

const chart: ChartPptxElement = {
	id: 'chart-1',
	type: 'chart',
	x: 0,
	y: 0,
	width: 400,
	height: 300,
	chartData: {
		chartType: 'bar',
		categories: ['A', 'B'],
		series: [{ name: 'S1', values: [1, 2], color: '#4472C4' }],
		style: { hasLegend: true },
	},
};

function setup(selected: PptxElement = chart) {
	const store = createStore<ViewerState>({
		...createInitialViewerState(),
		editable: true,
		slides: [{ id: 's1', elements: [selected] } as unknown as ViewerState['slides'][number]],
		selectedElementId: selected.id,
		selectedElementIds: [selected.id],
	});
	const patches: Array<Partial<PptxElement>> = [];
	const applyToSelected: ApplyToSelected = (build) => {
		const el = store.get().slides[0]!.elements[0]!;
		const patch = build(el);
		patches.push(patch);
		store.set({
			slides: [
				{ ...store.get().slides[0]!, elements: [{ ...el, ...patch } as PptxElement] },
			] as ViewerState['slides'],
		});
	};
	const actions = createChartFormatActions({ store, applyToSelected });
	const root = document.createElement('div');
	root.dataset.pptxEditorChrome = '';
	const body = document.createElement('div');
	body.dataset.pptxChrome = 'body';
	root.append(body);
	document.body.append(root);
	const host = createChartFormatPaneHost({
		doc: document,
		store,
		getTranslator: () => createTranslator(),
		actions,
		selectedElement: (state) => state.slides[0]?.elements[0],
	});
	host.attach(root);
	const unsubscribe = store.subscribe(() => host.sync());
	return { store, actions, root, body, host, patches, unsubscribe };
}

afterEach(() => document.body.replaceChildren());

describe('vanilla chart Format pane host', () => {
	it('opens on a chart double-click request, on the clicked element', () => {
		const { store, root, body, unsubscribe } = setup();
		const bar = document.createElement('span');
		root.append(bar);
		bar.dispatchEvent(
			new CustomEvent(CHART_FORMAT_REQUEST_EVENT, {
				bubbles: true,
				detail: { elementId: 'chart-1', target: 'legend', open: true },
			}),
		);
		const pane = body.querySelector('pptx-ui-chart-format-pane');
		expect(pane).not.toBeNull();
		expect(pane!.target).toBe('legend');
		expect(chartFormatTargetOf(store.get())).toBe('legend');
		unsubscribe();
	});

	it('applies a committed pane edit to the selected chart and closes on request', () => {
		const { store, actions, body, patches, unsubscribe } = setup();
		actions.openChartFormatPane('series:0');
		const pane = body.querySelector('pptx-ui-chart-format-pane')!;
		const next = { ...chart.chartData!, barGapWidth: 80 };
		pane.dispatchEvent(
			new CustomEvent('chart-format-change', { detail: { chartData: next, target: 'series:0' } }),
		);
		expect(patches).toStrictEqual([{ chartData: next }]);
		pane.dispatchEvent(new CustomEvent('chart-format-close'));
		expect(store.get().chartFormatPaneOpen).toBeFalsy();
		expect(body.querySelector('pptx-ui-chart-format-pane')).toBeNull();
		unsubscribe();
	});

	it('resets (Reset to Match Style) the current element of the selected chart', () => {
		const { actions, patches, unsubscribe } = setup();
		actions.selectChartFormatTarget('series:0');
		actions.resetChartFormatSelection();
		const data = (patches[0] as Partial<ChartPptxElement>).chartData!;
		expect(data.series[0]!.color).toBeUndefined();
		expect(data.elementFormats?.['series:0']).toStrictEqual({ reset: true });
		unsubscribe();
	});

	it('ignores requests for another element and shows nothing for a non-chart', async () => {
		const shape = {
			id: 'shape-1',
			type: 'shape',
			x: 0,
			y: 0,
			width: 10,
			height: 10,
		} as PptxElement;
		const { store, actions, root, body, unsubscribe } = setup(shape);
		root.dispatchEvent(
			new CustomEvent(CHART_FORMAT_REQUEST_EVENT, {
				detail: { elementId: 'chart-1', target: 'legend', open: true },
			}),
		);
		expect(store.get().chartFormatPaneOpen).toBeFalsy();
		actions.openChartFormatPane();
		expect(body.querySelector('pptx-ui-chart-format-pane')).toBeNull();
		await Promise.resolve();
		expect(store.get().chartFormatPaneOpen).toBeFalsy();
		unsubscribe();
	});
});
