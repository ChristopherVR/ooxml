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
	it('disables ambiguous multi-trendline controls while preserving all their source settings', () => {
		const space = columnSpace();
		space.plotArea.groups[0]!.series[0]!.trendlines = [
			{ trendlineType: 'linear' },
			{ trendlineType: 'poly', order: 3 },
		];
		const source = chartSpaceFormatSource(space);
		const model = buildChartFormatModel(source, 'series:0', 'options');
		expect(model.sections.find((section) => section.id === 'trendlines')?.fields[0]?.disabled).toBe(
			true,
		);
		expect(
			chartFormatFieldIntent(source.element(model.target), model.target, 'trendline.type', 'exp'),
		).toBeNull();
		expect(space.plotArea.groups[0]!.series[0]!.trendlines).toHaveLength(2);
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
});
