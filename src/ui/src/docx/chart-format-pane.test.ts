// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { history, undo, redo } from 'prosemirror-history';
import { EditorState, NodeSelection, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import type { ChartEditIntent } from 'ooxml-core/chart';
import {
	chartsIn,
	createDocument,
	loadDocx,
	newChartImage,
	newDocxChart,
	saveDocx,
} from 'ooxml-core/docx';
import { insertChartImage, selectedChart } from 'ooxml-core/docx/ui';
import { createChartFormatPane } from './chart-format-pane';
import { computeMenuItems, menuStateFromView } from './context-menu-items';
import { resizePicture } from './image-media';
import type { EditorLocale } from './localization';
import { CHART_FORMAT_STRINGS } from '../chart/format-pane-locales/index';
import { docToModel, modelToDoc } from './model-adapter';

afterEach(() => {
	document.body.replaceChildren();
});

function editorWithChart() {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Hello' }] }];
	const doc = modelToDoc(model);
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, {
		state: EditorState.create({
			doc,
			selection: TextSelection.create(doc, 6),
			plugins: [history()],
		}),
	});
	view.dispatch(
		insertChartImage(view.state, newChartImage(newDocxChart(undefined, { title: 'Sales' })))!,
	);
	return { model, view };
}

function paneFor(view: EditorView, locale: EditorLocale = 'en') {
	const pane = createChartFormatPane({
		view: () => view,
		theme: () => undefined,
		locale: () => locale,
		canEdit: () => true,
	});
	document.body.append(pane.element);
	return pane;
}

function edit(pane: ReturnType<typeof paneFor>, intent: ChartEditIntent, field = 'fill.color') {
	pane.element.dispatchEvent(
		new CustomEvent('office-chart-format-edit', {
			detail: { target: 'chartArea', field, value: '', intent },
		}),
	);
}

const RED: ChartEditIntent = {
	op: 'fill',
	target: { kind: 'series', seriesIndex: 0 },
	fill: { kind: 'solid', color: { kind: 'srgb', value: 'FF0000', transforms: [] } },
};

describe('Word chart Format pane', () => {
	it('saves advanced shared intents and undoes the last analysis edit independently', async () => {
		const { view, model } = editorWithChart();
		const pane = paneFor(view);
		pane.open('series:0');
		try {
			edit(
				pane,
				{ op: 'trendline', seriesIndex: 0, index: 0, edit: { type: 'poly', order: 3 } },
				'trendline.type',
			);
			edit(
				pane,
				{
					op: 'errorBars',
					seriesIndex: 0,
					direction: 'y',
					edit: { valueType: 'fixedVal', value: 2 },
				},
				'errorBars.type',
			);
			undo(view.state, view.dispatch);
			expect(
				selectedChart(view.state)?.chart.chartSpace?.plotArea.groups[0]!.series[0]!.errorBars
					?.length ?? 0,
			).toBe(0);
			expect(
				selectedChart(view.state)?.chart.chartSpace?.plotArea.groups[0]!.series[0]!.trendlines?.[0]
					?.order,
			).toBe(3);
			redo(view.state, view.dispatch);
			const bytes = await saveDocx(docToModel(view.state.doc, model));
			const xml = await (
				await JSZip.loadAsync(bytes)
			)
				.file('word/charts/chart1.xml')!
				.async('string');
			expect(xml).toContain('<c:trendlineType val="poly"/>');
			expect(xml).toContain('<c:order val="3"/>');
			expect(xml).toContain('<c:errBars>');
			const reloaded = await loadDocx(bytes);
			expect(chartsIn(reloaded.model.blocks)).toHaveLength(1);
		} finally {
			view.destroy();
		}
	});
	it('opens for the selected chart and applies each intent as one undo step', async () => {
		const { view, model } = editorWithChart();
		const pane = paneFor(view);
		pane.open('series:0');
		expect(pane.isOpen).toBe(true);
		expect(pane.element.chartSpace).toEqual(selectedChart(view.state)?.chart.chartSpace);
		expect(pane.element.disabled).toBe(false);
		const space = () => selectedChart(view.state)?.chart.chartSpace;
		const initial = space()?.legend?.position;
		edit(pane, RED);
		edit(pane, { op: 'legend', edit: { position: 't' } }, 'legend.position');
		expect(space()?.legend?.position).toBe('t');
		expect(pane.element.chartSpace).toEqual(space());
		undo(view.state, view.dispatch);
		expect(space()?.legend?.position).toBe(initial);
		expect(selectedChart(view.state)).toBeDefined();
		// Saved through the chart save path: the series fill is in the new chart part.
		const bytes = await saveDocx(docToModel(view.state.doc, model));
		const part = await (
			await JSZip.loadAsync(bytes)
		)
			.file('word/charts/chart1.xml')!
			.async('string');
		expect(part).toContain('<a:srgbClr val="FF0000"/>');
	});

	it('ignores fields without an intent and closes when the chart is deselected', () => {
		const { view } = editorWithChart();
		const pane = paneFor(view);
		pane.open();
		const before = view.state.doc;
		pane.element.dispatchEvent(
			new CustomEvent('office-chart-format-edit', {
				detail: { target: 'chartArea', field: 'effects.shadow', value: true, intent: null },
			}),
		);
		expect(view.state.doc).toBe(before);
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
		pane.refresh();
		expect(pane.isOpen).toBe(false);
	});

	it('shows every caption from the pane dictionary in the editor locale', () => {
		const { view } = editorWithChart();
		const pane = paneFor(view, 'fr');
		const seen = new Set<string>();
		pane.open();
		const translate = pane.element.translateLabel;
		pane.element.translateLabel = (key, params, fallback) => {
			seen.add(key);
			return translate(key, params, fallback);
		};
		const targets = [
			...pane.element.shadowRoot!.querySelectorAll('[data-chart-format-target] option'),
		];
		expect(targets.length).toBeGreaterThan(4);
		for (const option of targets) {
			pane.element.target = (option as HTMLOptionElement).value;
			for (const tab of pane.element.shadowRoot!.querySelectorAll<HTMLElement>(
				'[data-chart-format-tab]',
			))
				tab.click();
		}
		const missing = [...seen].filter((key) => !(key in CHART_FORMAT_STRINGS.fr));
		expect(missing).toEqual([]);
		expect(pane.element.getAttribute('aria-label')).not.toMatch(/^Format/);
	});

	it('offers Edit Data and Format Chart Area on the context menu of a selected chart', () => {
		const { view } = editorWithChart();
		const ids = computeMenuItems(menuStateFromView(view, false)).map((item) => item.id);
		expect(ids).toContain('chart-format');
		expect(ids).toContain('chart-data');
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
		expect(computeMenuItems(menuStateFromView(view, false)).map((item) => item.id)).not.toContain(
			'chart-format',
		);
	});

	it('resizes a chart as one undo step and saves the new extent', async () => {
		const { view, model } = editorWithChart();
		const pos = selectedChart(view.state)!.pos;
		resizePicture(view, pos, 400, 200);
		expect(view.state.selection).toBeInstanceOf(NodeSelection);
		const node = view.state.doc.nodeAt(pos)!;
		expect([node.attrs.widthPx, node.attrs.heightPx]).toEqual([400, 200]);
		const bytes = await saveDocx(docToModel(view.state.doc, model));
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toContain('<wp:extent cx="3810000" cy="1905000"/>');
		const reloaded = await loadDocx(bytes);
		expect(chartsIn(reloaded.model.blocks)).toHaveLength(1);
		undo(view.state, view.dispatch);
		expect(view.state.doc.nodeAt(pos)?.attrs.widthPx).toBe(576);
	});
});
