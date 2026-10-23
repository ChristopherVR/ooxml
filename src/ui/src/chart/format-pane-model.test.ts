import { applyChartEditIntent, type ChartSpace } from 'ooxml-core/chart';
import { describe, expect, it } from 'vitest';

import { translateChartFormatText } from './format-pane-common';
import { columnSpace } from './format-pane-fixture.test-util';
import { chartFormatFieldIntent } from './format-pane-intent';
import { buildChartFormatModel, chartFormatSourceTabs } from './format-pane-model';
import { chartSpaceFormatSource } from './format-pane-space';
import type { ChartFormatFieldValue } from './format-pane-types';

const english = (text: Parameters<typeof translateChartFormatText>[0]) =>
	translateChartFormatText(text, (key) => key);

/** Applies field `field` = `value` on `target` the way a host of the pane does. */
function edit(space: ChartSpace, target: string, field: string, value: ChartFormatFieldValue) {
	const source = chartSpaceFormatSource(space);
	const model = buildChartFormatModel(source, target);
	const intent = chartFormatFieldIntent(source.element(model.target), model.target, field, value);
	if (!intent) throw new Error(`No intent for ${field}`);
	return applyChartEditIntent(space, intent);
}

describe('chartSpaceFormatSource', () => {
	it('lists the chart elements in picker order with English captions', () => {
		const source = chartSpaceFormatSource(columnSpace());
		expect(source.targets.map((option) => english(option.label))).toEqual([
			'Chart Area',
			'Chart Title',
			'Horizontal (Category) Axis',
			'Vertical (Value) Axis',
			'Vertical (Value) Axis Major Gridlines',
			'Legend',
			'Plot Area',
			'Series "East"',
			'Series "East" Data Labels',
			'Series "West"',
			'Series "West" Data Labels',
		]);
	});

	it('offers only what the neutral operations write', () => {
		const source = chartSpaceFormatSource(columnSpace());
		const tabs = (key: string) =>
			chartFormatSourceTabs(source, buildChartFormatModel(source, key).target).map((t) => t.id);
		expect(tabs('chartArea')).toEqual(['fillLine']);
		expect(tabs('series:0')).toEqual(['fillLine', 'options']);
		expect(tabs('axis:1')).toEqual(['fillLine', 'options']);
		const model = buildChartFormatModel(source, 'series:0');
		const fill = model.sections.find((section) => section.id === 'fill')!;
		const kinds = fill.fields[0]!;
		expect(kinds.kind === 'radio' && kinds.choices.map((c) => c.value)).toEqual([
			'none',
			'solid',
			'gradient',
			'automatic',
		]);
		expect(fill.fields.find((f) => f.id === 'fill.color')).toMatchObject({ value: '#4472C4' });
		const line = model.sections.find((section) => section.id === 'line')!;
		expect(line.fields.map((f) => f.id)).not.toContain('line.compound');
		const options = buildChartFormatModel(source, 'series:0', 'options');
		expect(options.sections[0]!.fields.map((f) => [f.id, f.kind === 'number' && f.value])).toEqual([
			['series.overlap', -27],
			['series.gapWidth', 219],
		]);
	});

	it('falls back to the chart area and puts every section on one page for a flat source', () => {
		const source = {
			...chartSpaceFormatSource(columnSpace()),
			flat: true,
			slots: { fill: 'fill' },
		};
		expect(buildChartFormatModel(source, 'series:9').targetKey).toBe('chartArea');
		const model = buildChartFormatModel(source, 'series:1');
		expect(model.tabs).toEqual([]);
		expect(model.sections.map((s) => [s.id, s.slot])).toEqual([
			['fill', 'fill'],
			['line', undefined],
			['seriesOptions', undefined],
		]);
	});
});

describe('chartFormatFieldIntent', () => {
	it('turns paint fields into fill, line and transparency edits', () => {
		let space = edit(columnSpace(), 'plotArea', 'fill.kind', 'solid');
		space = edit(space, 'plotArea', 'fill.color', '#112233');
		space = edit(space, 'plotArea', 'fill.transparency', 40);
		expect(space.plotArea.spPr?.fill).toEqual({
			kind: 'solid',
			color: { kind: 'srgb', value: '112233', transforms: [{ name: 'alpha', value: '60000' }] },
		});
		space = edit(space, 'series:1', 'line.kind', 'solid');
		space = edit(space, 'series:1', 'line.width', 2);
		space = edit(space, 'series:1', 'line.dash', 'dash');
		expect(space.plotArea.groups[0]!.series[1]!.spPr?.line).toMatchObject({
			widthEmu: 25400,
			dash: 'dash',
			fill: { kind: 'solid' },
		});
		space = edit(space, 'chartArea', 'fill.kind', 'gradient');
		expect(space.spPr?.fill).toMatchObject({ kind: 'gradient', angle: 90 });
		space = edit(space, 'chartArea', 'fill.gradientType', 'radial');
		expect(space.spPr?.fill).toMatchObject({ kind: 'gradient', path: 'circle' });
		expect(edit(space, 'chartArea', 'fill.kind', 'automatic').spPr?.fill).toBeUndefined();
	});

	it('turns options fields into axis, legend, label and group edits', () => {
		let space = edit(columnSpace(), 'axis:1', 'axis.max', 50);
		space = edit(space, 'axis:1', 'axis.logScale', true);
		space = edit(space, 'axis:1', 'axis.reverseOrder', true);
		space = edit(space, 'axis:1', 'axis.displayUnits', 'thousands');
		expect(space.plotArea.axes[1]).toMatchObject({
			scaling: { max: 50, logBase: 10, orientation: 'maxMin' },
			displayUnits: { builtInUnit: 'thousands' },
		});
		expect(edit(space, 'axis:1', 'axis.max', '').plotArea.axes[1]!.scaling.max).toBeUndefined();
		space = edit(space, 'legend', 'legend.position', 'b');
		space = edit(space, 'legend', 'legend.noOverlap', false);
		expect(space.legend).toMatchObject({ position: 'b', overlay: true });
		space = edit(space, 'dataLabels:0', 'labels.showValue', true);
		expect(space.plotArea.groups[0]!.series[0]!.dataLabels?.showValue).toBe(true);
		space = edit(space, 'series:0', 'series.gapWidth', 75);
		expect(space.plotArea.groups[0]!.gapWidth).toBe(75);
	});

	it('refuses unusable values and fields without a neutral operation', () => {
		const space = columnSpace();
		const source = chartSpaceFormatSource(space);
		const view = source.element({ kind: 'series', seriesIndex: 0 });
		const series = { kind: 'series', seriesIndex: 0 } as const;
		expect(chartFormatFieldIntent(view, series, 'series.gapWidth', 900)).toBeNull();
		expect(chartFormatFieldIntent(view, series, 'fill.transparency', '')).toBeNull();
		expect(chartFormatFieldIntent(view, series, 'fill.kind', 'pattern')).toBeNull();
		expect(chartFormatFieldIntent(view, series, 'shadow.enabled', true)).toBeNull();
		expect(chartFormatFieldIntent(view, series, 'series.axis', 'secondary')).toBeNull();
	});
});
