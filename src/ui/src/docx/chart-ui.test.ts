// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { history, undo } from 'prosemirror-history';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { createDocument, saveDocx, loadDocx, chartsIn, type Paragraph } from 'ooxml-core/docx';
import { selectedChart } from 'ooxml-core/docx/ui';
import { docToModel, modelToDoc } from './model-adapter';
import { createChartDataDialog, createChartTypeDialog } from './chart-dialogs';
import { chartExFamilyOf } from 'ooxml-core/chart';
import { translate } from './localization';
import {
	chartPickerChoice,
	chartTypeTarget,
	setSelectedChartLegend,
	switchSelectedChartRowColumn,
	toggleSelectedChartElement,
} from './chart-commands';
import { syncChartDesign } from './chart-design-ribbon';
import { createRibbon } from './ribbon';
import { panelOf, isTabHidden } from './ribbon-tab-api';

function editor() {
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
	return { model, view };
}

type Picker = HTMLElement & { value: { type: string; grouping: string; title: string } };

function insertChart(view: EditorView, type = 'column', title = 'Sales') {
	const dialog = createChartTypeDialog({ view: () => view, theme: () => undefined }, false);
	document.body.append(dialog.element);
	dialog.open();
	const picker = dialog.element.querySelector('office-ui-chart-picker') as Picker;
	picker.value = { type, grouping: 'clustered', title };
	[...dialog.element.querySelectorAll('button')].find((b) => b.textContent === 'OK')!.click();
	expect(dialog.isOpen).toBe(false);
}

describe('Word charts in the editor', () => {
	it('maps picker choices onto the chart type switch and back', () => {
		expect(chartTypeTarget('bar', 'stacked')).toEqual({
			kind: 'bar',
			barDirection: 'bar',
			grouping: 'stacked',
		});
		expect(chartTypeTarget('pie', 'clustered')).toEqual({ kind: 'pie' });
	});

	it('inserts a chart from Insert Chart, selects it and shows Chart Design', () => {
		const { view, model } = editor();
		insertChart(view, 'bar');
		const selected = selectedChart(view.state)!;
		expect(selected.editable).toBe(true);
		expect(selected.chart.title).toBe('Sales');
		expect(chartPickerChoice(selected.chart)).toEqual({ type: 'bar', grouping: 'clustered' });
		const ribbon = createRibbon();
		syncChartDesign(ribbon, view.state, true);
		expect(isTabHidden(panelOf(ribbon, 'chart-design')!)).toBe(false);
		const run = (docToModel(view.state.doc, model).blocks[0] as Paragraph).runs.find(
			(item) => item.image?.chart,
		);
		expect(run?.image?.unsupported).toBe('Chart');
		undo(view.state, view.dispatch);
		expect(selectedChart(view.state)).toBeUndefined();
		syncChartDesign(ribbon, view.state, true);
		expect(isTabHidden(panelOf(ribbon, 'chart-design')!)).toBe(true);
	});

	it('inserts pareto and map charts from the extended families and saves them', async () => {
		for (const family of ['pareto', 'regionMap'] as const) {
			const { view, model } = editor();
			const dialog = createChartTypeDialog({ view: () => view, theme: () => undefined }, false);
			document.body.append(dialog.element);
			dialog.open();
			const picker = dialog.element.querySelector('office-ui-chart-picker') as Picker & {
				types: { id: string }[];
			};
			expect(picker.types.map((t) => t.id)).toContain(`cx:${family}`);
			picker.value = { type: `cx:${family}`, grouping: 'clustered', title: family };
			picker.dispatchEvent(new CustomEvent('office-chart-change'));
			const note = dialog.element.querySelector('[role="note"]')!.textContent;
			expect(note).toBe(family === 'regionMap' ? translate('en', 'chart.mapNotice') : '');
			[...dialog.element.querySelectorAll('button')].find((b) => b.textContent === 'OK')!.click();
			const selected = selectedChart(view.state)!;
			expect(selected.chart.chartExSpace).toBeDefined();
			const saved = await saveDocx(docToModel(view.state.doc, model));
			const chart = chartsIn((await loadDocx(saved)).model.blocks)[0]!;
			expect(chartExFamilyOf(chart.chartExSpace!)).toBe(family);
			expect(chart.title).toBe(family);
			dialog.element.remove();
		}
	});

	it('changes the type, the legend and the data of the selected chart, one undo step each', async () => {
		const { view, model } = editor();
		insertChart(view);
		const change = createChartTypeDialog({ view: () => view, theme: () => undefined }, true);
		change.open();
		(change.element.querySelector('office-ui-chart-picker') as Picker).value = {
			type: 'line',
			grouping: 'clustered',
			title: '',
		};
		[...change.element.querySelectorAll('button')].find((b) => b.textContent === 'OK')!.click();
		expect(selectedChart(view.state)?.chart.chartSpace?.plotArea.groups[0]?.kind).toBe('line');
		expect(setSelectedChartLegend(view, 'none')).toBe(true);
		expect(selectedChart(view.state)?.chart.chartSpace?.legend).toBeUndefined();
		const data = createChartDataDialog({ view: () => view, theme: () => undefined });
		data.open();
		const sheet = data.element.querySelector('office-ui-chart-data-sheet')!;
		sheet.dispatchEvent(
			new CustomEvent('office-chart-data-change', {
				detail: { edits: [{ kind: 'value', series: 0, category: 0, value: 42 }] },
			}),
		);
		const values = selectedChart(view.state)?.chart.chartSpace?.plotArea.groups[0]?.series[0]
			?.values?.cache?.points;
		expect(values?.[0]?.value).toBe('42');
		undo(view.state, view.dispatch);
		expect(
			selectedChart(view.state)?.chart.chartSpace?.plotArea.groups[0]?.series[0]?.values?.cache
				?.points[0]?.value,
		).toBe('4.3');
		// The inserted chart saves with its part and workbook and reloads as a line chart.
		const saved = await saveDocx(docToModel(view.state.doc, model));
		const chart = chartsIn((await loadDocx(saved)).model.blocks)[0]!;
		expect(chart.chartSpace?.plotArea.groups[0]?.kind).toBe('line');
		expect(chart.workbookPartName).toBe('word/embeddings/Microsoft_Excel_Worksheet.xlsx');
	});

	it('switches rows/columns and toggles chart elements via Chart Design commands', () => {
		const { view } = editor();
		insertChart(view);
		const beforeCats = selectedChart(view.state)?.chart.chartSpace?.plotArea.groups[0]?.series[0]
			?.categories?.cache?.points.length;
		const beforeSeries = selectedChart(view.state)?.chart.chartSpace?.plotArea.groups[0]?.series
			.length;

		expect(switchSelectedChartRowColumn(view)).toBe(true);
		const afterCats = selectedChart(view.state)?.chart.chartSpace?.plotArea.groups[0]?.series[0]
			?.categories?.cache?.points.length;
		const afterSeries = selectedChart(view.state)?.chart.chartSpace?.plotArea.groups[0]?.series
			.length;
		expect(afterCats).toBe(beforeSeries);
		expect(afterSeries).toBe(beforeCats);

		expect(toggleSelectedChartElement(view, 'dataLabels')).toBe(true);
		expect(
			selectedChart(view.state)?.chart.chartSpace?.plotArea.groups[0]?.dataLabels?.showValue,
		).toBe(true);
		expect(toggleSelectedChartElement(view, 'dataLabels')).toBe(true);
		expect(
			selectedChart(view.state)?.chart.chartSpace?.plotArea.groups[0]?.dataLabels?.showValue,
		).toBeFalsy();
	});
});
