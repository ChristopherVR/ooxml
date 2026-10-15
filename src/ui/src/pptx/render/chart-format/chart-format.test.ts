import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n';
import { buildContextMenuEntries } from '../context-menu-commands';
import type { ChartFormatText } from './chart-format-targets';
import {
	applyChartFormatField,
	buildChartFormatPaneModel,
	chartFormatTabs,
	chartSelectionHomeControls,
	dispatchChartSelectionRequest,
	listChartFormatTargets,
	resetChartFormatTarget,
	resolveChartFormatTarget,
	translateChartFormatText,
} from '.';
import { columnChart, doughnutChart, lineChart, pieChart } from './chart-format-fixtures.test-util';

const english = (key: string, params?: Readonly<Record<string, string | number>>) => {
	const value = translationsEn[key];
	return value
		? value.replace(/\{\{(\w+)\}\}/gu, (_, n: string) => String(params?.[n] ?? ''))
		: key;
};

describe('chart format targets', () => {
	it('lists the elements in PowerPoint picker order with English captions', () => {
		const labels = listChartFormatTargets(columnChart()).map((option) =>
			translateChartFormatText(option.label, english),
		);
		expect(labels).toEqual([
			'Chart Area',
			'Chart Title',
			'Horizontal (Category) Axis',
			'Vertical (Value) Axis',
			'Vertical (Value) Axis Major Gridlines',
			'Vertical (Value) Axis Minor Gridlines',
			'Secondary Vertical (Value) Axis',
			'Legend',
			'Plot Area',
			'Data Labels',
			'Series "East"',
			'Series "West"',
		]);
	});

	it('falls back to the chart area for a stale key', () => {
		expect(resolveChartFormatTarget(columnChart(), 'series:9').key).toBe('chartArea');
		expect(resolveChartFormatTarget(columnChart(), 'axis:2').key).toBe('axis:2');
		expect(resolveChartFormatTarget(undefined, 'legend').key).toBe('chartArea');
	});
});

/** Every caption a pane model can show, for the dictionary check. */
function captions(text: ChartFormatText | undefined, out: Set<string>): void {
	if (text) {
		out.add(text.key);
		for (const nested of Object.values(text.paramTexts ?? {})) {
			captions(nested, out);
		}
	}
}

describe('chart format pane model', () => {
	it('offers Options only where PowerPoint has them', () => {
		const tabs = (key: string) => chartFormatTabs(columnChart(), key).map((tab) => tab.id);
		expect(tabs('chartArea')).toEqual(['fillLine', 'effects']);
		expect(tabs('series:0')).toEqual(['fillLine', 'effects', 'options']);
		expect(tabs('axis:1')).toEqual(['fillLine', 'effects', 'options']);
		expect(tabs('legend')).toEqual(['fillLine', 'effects', 'options']);
	});

	it('shows Line only (no Fill) for gridlines and line series, with arrows', () => {
		const gridlines = buildChartFormatPaneModel(columnChart(), 'majorGridlines:1');
		expect(gridlines.title.fallback).toBe('Format Major Gridlines');
		expect(gridlines.sections.map((s) => s.id)).toEqual(['line']);
		const line = buildChartFormatPaneModel(lineChart(), 'series:0');
		expect(line.sections.map((s) => s.id)).toEqual(['line']);
		expect(line.sections[0]!.fields.map((f) => f.id)).toContain('line.endArrow');
	});

	it('has an English caption for every string any pane can show', () => {
		const keys = new Set<string>();
		for (const chart of [columnChart(), lineChart(), pieChart(), doughnutChart()]) {
			for (const option of listChartFormatTargets(chart)) {
				captions(option.label, keys);
				for (const tab of chartFormatTabs(chart, option.key)) {
					captions(tab.label, keys);
					let data = chart;
					for (const [field, value] of [
						['fill.kind', 'solid'],
						['line.kind', 'solid'],
						['shadow.enabled', true],
						['glow.enabled', true],
						['axis.crosses', 'value'],
					] as const) {
						data = applyChartFormatField(data, option.key, field, value) ?? data;
					}
					for (const model of [
						buildChartFormatPaneModel(data, option.key, tab.id),
						buildChartFormatPaneModel(
							applyChartFormatField(data, option.key, 'fill.kind', 'gradient') ?? data,
							option.key,
							tab.id,
						),
					]) {
						captions(model.title, keys);
						for (const section of model.sections) {
							captions(section.title, keys);
							captions(section.note, keys);
							for (const field of section.fields) {
								captions(field.label, keys);
								captions(field.hint, keys);
								for (const choice of 'choices' in field ? field.choices : []) {
									captions(choice.label, keys);
									captions(choice.hint, keys);
								}
							}
						}
					}
				}
			}
		}
		expect([...keys].filter((key) => !(key in translationsEn))).toEqual([]);
	});
});

describe('applyChartFormatField', () => {
	it('applies fill, line and effect fields through the core operations', () => {
		let data = applyChartFormatField(columnChart(), 'legend', 'fill.kind', 'solid')!;
		data = applyChartFormatField(data, 'legend', 'fill.color', '#112233')!;
		data = applyChartFormatField(data, 'legend', 'fill.transparency', 40)!;
		data = applyChartFormatField(data, 'legend', 'line.width', 2.5)!;
		data = applyChartFormatField(data, 'legend', 'shadow.enabled', true)!;
		data = applyChartFormatField(data, 'legend', 'shadow.blur', 8)!;
		expect(data.elementFormats?.legend).toMatchObject({
			fill: { kind: 'solid', color: '#112233', transparency: 40 },
			line: { width: 2.5 },
			effects: { shadow: { blur: 8 } },
		});
		const fill = buildChartFormatPaneModel(data, 'legend').sections[0]!;
		expect(fill.fields.find((f) => f.id === 'fill.transparency')).toMatchObject({ value: 40 });
	});

	it('rejects unusable values instead of writing them', () => {
		expect(applyChartFormatField(columnChart(), 'series:0', 'series.gapWidth', 900)).toBeNull();
		expect(applyChartFormatField(columnChart(), 'axis:1', 'axis.majorUnit', -1)).toBeNull();
		expect(applyChartFormatField(columnChart(), 'legend', 'fill.transparency', '')).toBeNull();
	});

	it('applies the Options fields and clears an axis bound back to Auto', () => {
		let data = applyChartFormatField(columnChart(), 'series:0', 'series.gapWidth', 75)!;
		expect(data.barGapWidth).toBe(75);
		data = applyChartFormatField(data, 'axis:1', 'axis.max', 100)!;
		expect(data.axes?.[1]?.max).toBe(100);
		data = applyChartFormatField(data, 'axis:1', 'axis.max', '')!;
		expect(data.axes?.[1]?.max).toBeUndefined();
		data = applyChartFormatField(data, 'legend', 'legend.position', 'b')!;
		data = applyChartFormatField(data, 'legend', 'legend.noOverlap', true)!;
		expect(data.style).toMatchObject({ legendPosition: 'b', legendOverlay: false });
		data = applyChartFormatField(data, 'dataLabels', 'labels.showCategory', true)!;
		expect(data.style?.dataLabels?.showCategory).toBe(true);
	});

	it('resets an element to match the chart style', () => {
		const formatted = applyChartFormatField(columnChart(), 'series:1', 'fill.color', '#FF0000')!;
		const reset = resetChartFormatTarget(formatted, 'series:1');
		expect(reset.series[1]?.color).toBeUndefined();
		expect(reset.elementFormats?.['series:1']).toEqual({ reset: true });
	});
});

describe('Current Selection strip', () => {
	it('lists the elements and gates the buttons', () => {
		const controls = chartSelectionHomeControls({
			chartData: columnChart(),
			targetKey: 'legend',
			editable: false,
			translate: english,
		});
		const elements = controls['chartFormat.currentSelection.chartElements'];
		expect(elements?.value).toBe('legend');
		expect(elements?.items?.[0]).toEqual({ value: 'chartArea', label: 'Chart Area' });
		expect(controls['chartFormat.currentSelection.formatSelection']?.disabled).toBe(false);
		expect(controls['chartFormat.currentSelection.resetToMatchStyle']?.disabled).toBe(true);
	});

	it('routes each request', () => {
		const calls: string[] = [];
		const handlers = {
			select: (key: string) => calls.push(`select:${key}`),
			formatSelection: () => calls.push('format'),
			reset: () => calls.push('reset'),
		};
		dispatchChartSelectionRequest(
			{ id: 'chartFormat.currentSelection.chartElements', value: 'axis:0' },
			handlers,
		);
		dispatchChartSelectionRequest({ id: 'chartFormat.currentSelection.formatSelection' }, handlers);
		dispatchChartSelectionRequest(
			{ id: 'chartFormat.currentSelection.resetToMatchStyle' },
			handlers,
		);
		expect(calls).toEqual(['select:axis:0', 'format', 'reset']);
	});
});

describe('chart context menu entry', () => {
	it('names the format entry after the clicked chart element', () => {
		const entry = (target?: string) =>
			buildContextMenuEntries({ elementType: 'chart', chartFormatTarget: target ?? null }).find(
				(e) => e.id === 'format-chart-element',
			);
		expect(translationsEn[entry()!.labelKey]).toBe('Format Chart Area...');
		expect(translationsEn[entry('series:0')!.labelKey]).toBe('Format Data Series...');
		expect(translationsEn[entry('axis:1')!.labelKey]).toBe('Format Axis...');
		expect(
			buildContextMenuEntries({ elementType: 'shape' }).some(
				(e) => e.id === 'format-chart-element',
			),
		).toBe(false);
	});
});
