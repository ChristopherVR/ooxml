import { afterEach, describe, expect, it } from 'vitest';
import {
	applyChartEditIntent,
	parseChartSpace,
	writeChartSpace,
	setChartType,
	type ChartSpace,
} from 'ooxml-core/chart';
import { columnSpace } from './format-pane-fixture.test-util';
import { chartSpaceFormatSource } from './format-pane-space';
import { buildChartFormatModel } from './format-pane-model';
import { chartFormatFieldIntent } from './format-pane-intent';
import { defineChartFormatPane, type ChartFormatEditEvent } from './format-pane';
import type { ChartFormatFieldValue } from './format-pane-types';

defineChartFormatPane();
afterEach(() => document.body.replaceChildren());
function edit(
	space: ChartSpace,
	key: string,
	field: string,
	value: ChartFormatFieldValue,
): ChartSpace {
	const source = chartSpaceFormatSource(space);
	const model = buildChartFormatModel(source, key);
	const intent = chartFormatFieldIntent(source.element(model.target), model.target, field, value);
	if (!intent) throw new Error(`Missing ${field} intent`);
	return applyChartEditIntent(space, intent);
}
const saved = (space: ChartSpace) => parseChartSpace(writeChartSpace(space)).chartSpace;

describe('advanced shared chart pane', () => {
	it('edits trendline and error-bar settings through actual supported field intents', () => {
		let space = edit(columnSpace(), 'series:0', 'trendline.type', 'poly');
		space = edit(space, 'series:0', 'trendline.order', 3);
		space = edit(space, 'series:0', 'trendline.equation', true);
		space = edit(space, 'series:0', 'errorBars.type', 'percentage');
		space = edit(space, 'series:0', 'errorBars.value', 7);
		space = edit(space, 'series:0', 'errorBars.cap', false);
		const first = saved(space).plotArea.groups[0]!.series[0]!;
		expect(first.trendlines?.[0]).toMatchObject({
			trendlineType: 'poly',
			order: 3,
			displayEquation: true,
		});
		expect(first.errorBars?.[0]).toMatchObject({
			valueType: 'percentage',
			value: 7,
			noEndCap: true,
		});
	});
	it('projects marker edits for line charts and keeps unsupported families free of analysis controls', () => {
		let space = edit(
			setChartType(columnSpace(), { kind: 'line' }),
			'series:0',
			'marker.symbol',
			'diamond',
		);
		space = edit(space, 'series:0', 'marker.size', 9);
		expect(saved(space).plotArea.groups[0]!.series[0]!.marker).toMatchObject({
			symbol: 'diamond',
			size: 9,
		});
		const pie = chartSpaceFormatSource(setChartType(space, { kind: 'pie' }));
		const sections = buildChartFormatModel(pie, 'series:0', 'options').sections.map(
			(section) => section.id,
		);
		expect(sections).not.toContain('markers');
		expect(sections).not.toContain('trendlines');
		expect(sections).not.toContain('errorBars');
	});
	it('gives each trendline its own section and edits only the one it shows', () => {
		const space = columnSpace();
		space.plotArea.groups[0]!.series[0]!.trendlines = [
			{ trendlineType: 'linear' },
			{ trendlineType: 'poly', order: 3 },
		];
		const model = buildChartFormatModel(chartSpaceFormatSource(space), 'series:0', 'options');
		const second = model.sections.find((section) => section.id === 'trendline-1');
		expect(second?.title).toMatchObject({
			key: 'chartFormat.section.trendlineN',
			params: { n: 2 },
		});
		expect(second?.fields.map((field) => field.id)).toContain('trendline.1.order');
		expect(second?.fields.at(-1)?.id).toBe('trendline.add');
		let next = edit(space, 'series:0', 'trendline.1.order', 5);
		next = edit(next, 'series:0', 'trendline.add', 'movingAvg');
		next = edit(next, 'series:0', 'trendline.2.period', 4);
		next = edit(next, 'series:0', 'trendline.type', 'exp');
		expect(
			saved(next).plotArea.groups[0]!.series[0]!.trendlines!.map((trend) => [
				trend.trendlineType,
				trend.order ?? trend.period,
			]),
		).toEqual([
			['exp', undefined],
			['poly', 5],
			['movingAvg', 4],
		]);
		next = edit(next, 'series:0', 'trendline.1.type', 'none');
		expect(saved(next).plotArea.groups[0]!.series[0]!.trendlines).toHaveLength(2);
		expect(space.plotArea.groups[0]!.series[0]!.trendlines).toHaveLength(2);
	});
	it('edits custom error-bar sources as values or cached ranges, refusing unknown cells', () => {
		const start = columnSpace();
		const range = 'Sheet1!$C$2:$C$3';
		const west = start.plotArea.groups[0]!.series[1]!;
		west.values = { ...west.values!, kind: 'numRef', formula: range };
		let space = edit(start, 'series:0', 'errorBars.type', 'cust');
		const bars = () => saved(space).plotArea.groups[0]!.series[0]!.errorBars![0]!;
		expect(bars().plus?.cache?.points.map((point) => point.value)).toEqual(['1', '1']);
		space = edit(space, 'series:0', 'errorBars.plusValues', '2, 3');
		space = edit(space, 'series:0', 'errorBars.minusValues', range);
		expect(bars()).toMatchObject({ valueType: 'cust', minus: { kind: 'numRef', formula: range } });
		expect(bars().minus?.cache?.points.map((point) => point.value)).toEqual(['1', '2']);
		expect(bars().plus?.cache?.points.map((point) => point.value)).toEqual(['2', '3']);
		const source = chartSpaceFormatSource(space);
		const model = buildChartFormatModel(source, 'series:0', 'options');
		const section = model.sections.find((entry) => entry.id === 'errorBars')!;
		expect(section.fields.find((field) => field.id === 'errorBars.minusValues')?.value).toBe(range);
		expect(section.note?.key).toBe('chartFormat.errorBars.customCachedHint');
		const view = source.element(model.target);
		for (const bad of ['Sheet9!$Z$1:$Z$3', '-1', 'abc'])
			expect(chartFormatFieldIntent(view, model.target, 'errorBars.plusValues', bad)).toBeNull();
		// A workbook resolver admits ranges the chart does not cache.
		const workbook = chartSpaceFormatSource(space, { resolveRange: () => [5, 6, 7] });
		const intent = chartFormatFieldIntent(
			workbook.element(model.target),
			model.target,
			'errorBars.plusValues',
			'Sheet9!$Z$1:$Z$3',
		);
		expect(intent).toMatchObject({
			op: 'errorBars',
			edit: { plus: { formula: 'Sheet9!$Z$1:$Z$3' } },
		});
	});
	it('edits source-backed shape effects, bevels, patterns and line resets with no source mutation', () => {
		const source = columnSpace();
		const before = structuredClone(source);
		let space = edit(source, 'chartArea', 'shadow.enabled', true);
		space = edit(space, 'chartArea', 'shadow.blur', 6);
		space = edit(space, 'chartArea', 'threeD.bevelTop', 'circle');
		space = edit(space, 'chartArea', 'fill.kind', 'pattern');
		space = edit(space, 'chartArea', 'fill.patternPreset', 'cross');
		space = edit(space, 'chartArea', 'fill.patternForeground', '#123456');
		space = edit(space, 'axis:1', 'line.compound', 'dbl');
		space = edit(space, 'axis:1', 'line.beginArrow', 'triangle');
		space = edit(space, 'axis:1', 'line.width', 2);
		space = edit(space, 'axis:1', 'line.width', '');
		const result = saved(space);
		expect(result.spPr?.fill).toMatchObject({
			kind: 'pattern',
			preset: 'cross',
			foreground: { value: '123456' },
		});
		expect(result.spPr?.effectsXml).toContain('blurRad="76200"');
		expect(result.spPr?.sourceXml).toContain('bevelT');
		expect(result.plotArea.axes[1]!.spPr?.line).toMatchObject({
			compound: 'dbl',
			headEnd: { type: 'triangle' },
		});
		expect(result.plotArea.axes[1]!.spPr?.line?.widthEmu).toBeUndefined();
		expect(source).toEqual(before);
	});
	it('edits title text geometry and an axis title without changing displayed text', () => {
		let space = edit(columnSpace(), 'title', 'text.fill.kind', 'solid');
		space = edit(space, 'title', 'text.fill.color', '#123456');
		space = edit(space, 'title', 'text.rotation', 25);
		space = edit(space, 'title', 'text.glow.enabled', true);
		space = edit(space, 'axis:1', 'axis.title', 'Revenue');
		space = edit(space, 'axis:1', 'axis.numberFormat', '0.0');
		space = edit(space, 'axis:1', 'axis.numberFormatSourceLinked', false);
		const result = saved(space);
		expect(result.title?.text).toBe(columnSpace().title?.text);
		expect(result.title?.tx?.rich?.sourceXml).toContain('rot="1500000"');
		expect(result.plotArea.axes[1]?.title?.text).toBe('Revenue');
		expect(result.plotArea.axes[1]?.numberFormat?.sourceLinked).toBe(false);
	});
	it('retains basic-only capability bounds for narrow hosts', () => {
		const source = chartSpaceFormatSource(columnSpace(), { advanced: false });
		const target = { kind: 'chartArea' } as const;
		const view = source.element(target);
		expect(view.support.effects).toBe(false);
		expect(view.support.pattern).toBe(false);
		expect(source.textOptions).toBe(false);
		expect(chartFormatFieldIntent(view, target, 'shadow.enabled', true)).toBeNull();
		expect(chartFormatFieldIntent(view, target, 'fill.kind', 'pattern')).toBeNull();
	});
	it('emits a usable analysis intent from the rendered pane and redraws its dependent fields', () => {
		const pane = document.createElement('office-ui-chart-format-pane');
		pane.chartSpace = columnSpace();
		pane.target = 'series:0';
		pane.addEventListener('office-chart-format-edit', (event) => {
			const intent = (event as ChartFormatEditEvent).detail.intent;
			if (intent) pane.chartSpace = applyChartEditIntent(pane.chartSpace!, intent);
		});
		document.body.append(pane);
		pane.shadowRoot!.querySelector<HTMLButtonElement>('[data-chart-format-tab="options"]')!.click();
		const select = pane.shadowRoot!.querySelector<HTMLSelectElement>(
			'[data-chart-format-field="trendline.type"] select',
		)!;
		select.value = 'poly';
		select.dispatchEvent(new Event('change'));
		expect(pane.chartSpace!.plotArea.groups[0]!.series[0]!.trendlines?.[0]?.trendlineType).toBe(
			'poly',
		);
		expect(
			pane.shadowRoot!.querySelector('[data-chart-format-field="trendline.order"]'),
		).not.toBeNull();
	});
	it('adds a second trendline and types a custom error source in the rendered pane', () => {
		const pane = document.createElement('office-ui-chart-format-pane');
		pane.chartSpace = columnSpace();
		pane.target = 'series:0';
		pane.addEventListener('office-chart-format-edit', (event) => {
			const intent = (event as ChartFormatEditEvent).detail.intent;
			if (intent) pane.chartSpace = applyChartEditIntent(pane.chartSpace!, intent);
		});
		document.body.append(pane);
		const root = pane.shadowRoot!;
		root.querySelector<HTMLButtonElement>('[data-chart-format-tab="options"]')!.click();
		const change = (field: string, tag: 'select' | 'input', value: string) => {
			const control = root.querySelector<HTMLInputElement | HTMLSelectElement>(
				`[data-chart-format-field="${field}"] ${tag}`,
			)!;
			control.value = value;
			control.dispatchEvent(new Event('change'));
		};
		change('trendline.type', 'select', 'linear');
		change('trendline.add', 'select', 'movingAvg');
		expect(root.querySelector('[data-chart-format-field="trendline.1.period"]')).not.toBeNull();
		change('errorBars.type', 'select', 'cust');
		change('errorBars.plusValues', 'input', '0.5, 1.5');
		const series = pane.chartSpace!.plotArea.groups[0]!.series[0]!;
		expect(series.trendlines?.map((trend) => trend.trendlineType)).toEqual(['linear', 'movingAvg']);
		expect(series.errorBars?.[0]?.plus?.cache?.points.map((point) => point.value)).toEqual([
			'0.5',
			'1.5',
		]);
	});
});
