// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { expect, it } from 'vitest';
import { history, undo } from 'prosemirror-history';
import { EditorState, NodeSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { loadDocx } from 'ooxml-core/docx';
import { selectedChart } from 'ooxml-core/docx/ui';
import { modelToDoc } from './model-adapter';
import { createChartDataDialog, createChartTypeDialog } from './chart-dialogs';
import { createChartFormatPane } from './chart-format-pane';

it('edits a ChartEx title, family options and formatting through shared controls and history', async () => {
	const loaded = await loadDocx(
		readFileSync(
			path.join(import.meta.dirname, '../../../core/docx/__fixtures__/chart-families.docx'),
		),
	);
	const doc = modelToDoc(loaded.model);
	let position = -1;
	doc.descendants((node, pos) => {
		if (
			position < 0 &&
			node.type.name === 'image' &&
			JSON.parse(String(node.attrs.chart ?? 'null'))?.layoutId === 'waterfall'
		)
			position = pos;
	});
	expect(position).toBeGreaterThanOrEqual(0);
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, {
		state: EditorState.create({
			doc,
			selection: NodeSelection.create(doc, position),
			plugins: [history()],
		}),
	});
	try {
		expect(selectedChart(view.state)?.editable).toBe(true);
		const dialog = createChartDataDialog({ view: () => view, theme: () => undefined });
		document.body.append(dialog.element);
		dialog.open();
		const options = dialog.element.querySelector('office-ui-chart-ex-options')!;
		options.dispatchEvent(
			new CustomEvent('office-chart-ex-edit', {
				detail: { action: { op: 'title', text: 'New title' } },
			}),
		);
		expect(selectedChart(view.state)?.chart.title).toBe('New title');
		undo(view.state, view.dispatch);
		expect(selectedChart(view.state)?.chart.title).not.toBe('New title');
		options.dispatchEvent(
			new CustomEvent('office-chart-ex-edit', {
				detail: { action: { op: 'waterfall', seriesIndex: 0, options: { connectorLines: false } } },
			}),
		);
		expect(
			selectedChart(view.state)?.chart.chartExSpace?.plotArea.series[0]!.layoutProperties
				?.visibility?.connectorLines,
		).toBe(false);
		dialog.close();
		const change = createChartTypeDialog({ view: () => view, theme: () => undefined }, true);
		document.body.append(change.element);
		change.open();
		change.element.querySelector('office-ui-chart-ex-options')!.dispatchEvent(
			new CustomEvent('office-chart-ex-edit', {
				detail: { action: { op: 'layout', layout: 'treemap' } },
			}),
		);
		expect(selectedChart(view.state)?.chart.layoutId).toBe('treemap');
		change.close();
		undo(view.state, view.dispatch);
		const pane = createChartFormatPane({
			view: () => view,
			theme: () => undefined,
			locale: () => 'en',
			canEdit: () => true,
		});
		document.body.append(pane.element);
		pane.open();
		expect(pane.isOpen).toBe(true);
		pane.element.dispatchEvent(
			new CustomEvent('office-chart-format-edit', {
				detail: {
					intent: {
						op: 'fill',
						target: { kind: 'series', seriesIndex: 0 },
						fill: { kind: 'solid', color: { kind: 'srgb', value: 'FF0000', transforms: [] } },
					},
				},
			}),
		);
		expect(
			selectedChart(view.state)?.chart.chartExSpace?.plotArea.series[0]?.spPr?.fill,
		).toMatchObject({ kind: 'solid', color: { value: 'FF0000' } });
	} finally {
		view.destroy();
		document.body.replaceChildren();
	}
});
