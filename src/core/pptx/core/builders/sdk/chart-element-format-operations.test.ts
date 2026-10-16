import { describe, expect, it } from 'vitest';

import type { PptxChartData } from '../../types';
import {
	chartFormatTargetKey,
	parseChartFormatTargetKey,
} from '../../utils/chart-element-format-keys';
import {
	readChartElementFormat,
	resetChartElementFormat,
	withChartElementFormat,
} from './chart-element-format-operations';
import {
	withChartAxisOptions,
	withChartDataLabelOptions,
	withChartLegendOptions,
	withChartSeriesOptions,
} from './chart-format-options-operations';

function bar(): PptxChartData {
	return {
		chartType: 'bar',
		categories: ['A', 'B'],
		series: [{ name: 'S1', values: [1, 2], color: '#4472C4' }],
		style: { chartAreaFill: '#FFFFFF', hasLegend: true },
		axes: [
			{ axisType: 'catAx', axisId: 1 },
			{ axisType: 'valAx', axisId: 2, majorGridlines: true },
		],
	};
}

describe('chart format target keys', () => {
	it('round-trips every target kind', () => {
		for (const key of [
			'chartArea',
			'plotArea',
			'title',
			'legend',
			'axis:1',
			'majorGridlines:0',
			'minorGridlines:2',
			'series:3',
			'dataLabels',
			'dataLabels:1',
		]) {
			const target = parseChartFormatTargetKey(key);
			expect(target).toBeDefined();
			expect(chartFormatTargetKey(target!)).toBe(key);
		}
		expect(parseChartFormatTargetKey('axis')).toBeUndefined();
		expect(parseChartFormatTargetKey('series:-1')).toBeUndefined();
		expect(parseChartFormatTargetKey('nope')).toBeUndefined();
	});
});

describe('withChartElementFormat', () => {
	it('records the edit and mirrors the flat fields without mutating its input', () => {
		const data = bar();
		const next = withChartElementFormat(
			data,
			{ kind: 'chartArea' },
			{ fill: { kind: 'solid', color: '#FF0000', transparency: 30 }, line: { kind: 'none' } },
		);
		expect(data.style?.chartAreaFill).toBe('#FFFFFF');
		expect(data.elementFormats).toBeUndefined();
		expect(next.style?.chartAreaFill).toBe('#FF0000');
		expect(next.style?.chartAreaBorder).toBe('none');
		expect(next.elementFormats?.chartArea).toEqual({
			fill: { kind: 'solid', color: '#FF0000', transparency: 30 },
			line: { kind: 'none' },
		});
		const view = readChartElementFormat(next, { kind: 'chartArea' });
		expect(view.fill).toMatchObject({ kind: 'solid', color: '#FF0000', transparency: 30 });
		expect(view.line.kind).toBe('none');
	});

	it('keeps earlier sections when a later edit changes another one', () => {
		let next = withChartElementFormat(
			bar(),
			{ kind: 'legend' },
			{ fill: { kind: 'solid', color: '#00FF00' } },
		);
		next = withChartElementFormat(next, { kind: 'legend' }, { line: { width: 2 } });
		next = withChartElementFormat(next, { kind: 'legend' }, { fill: { transparency: 50 } });
		expect(next.elementFormats?.legend).toEqual({
			fill: { kind: 'solid', color: '#00FF00', transparency: 50 },
			line: { kind: 'automatic', width: 2 },
		});
	});

	it('writes a series fill to its colour and a gridline line to the gridline shape', () => {
		let next = withChartElementFormat(
			bar(),
			{ kind: 'series', seriesIndex: 0 },
			{ fill: { kind: 'solid', color: '#C00000' } },
		);
		next = withChartElementFormat(
			next,
			{ kind: 'majorGridlines', axisIndex: 1 },
			{ line: { kind: 'solid', color: '#D9D9D9', width: 0.5, dash: 'dash' } },
		);
		expect(next.series[0].color).toBe('#C00000');
		expect(next.axes?.[1].majorGridlinesSpPr).toEqual({
			strokeColor: '#D9D9D9',
			strokeWidth: 0.5,
			strokeDashStyle: 'dash',
		});
	});

	it('treats a line series as a Line-only target', () => {
		const data: PptxChartData = { ...bar(), chartType: 'line' };
		expect(readChartElementFormat(data, { kind: 'series', seriesIndex: 0 }).support.fill).toBe(
			false,
		);
		const next = withChartElementFormat(
			data,
			{ kind: 'series', seriesIndex: 0 },
			{ line: { kind: 'solid', color: '#00B050', width: 3 } },
		);
		expect(next.series[0]).toMatchObject({ color: '#00B050', lineWidth: 3 });
	});

	it('switches a chart area to a gradient through the gradient field', () => {
		const next = withChartElementFormat(
			bar(),
			{ kind: 'chartArea' },
			{ fill: { kind: 'gradient' } },
		);
		expect(next.style?.chartAreaGradient?.stops).toHaveLength(2);
		expect(next.style?.chartAreaFill).toBeUndefined();
		expect(readChartElementFormat(next, { kind: 'chartArea' }).fill.kind).toBe('gradient');
	});
});

describe('resetChartElementFormat', () => {
	it('clears flat overrides and records a reset', () => {
		const formatted = withChartElementFormat(
			bar(),
			{ kind: 'series', seriesIndex: 0 },
			{ line: { width: 4 } },
		);
		const next = resetChartElementFormat(formatted, { kind: 'series', seriesIndex: 0 });
		expect(next.series[0].color).toBeUndefined();
		expect(next.series[0].lineWidth).toBeUndefined();
		expect(next.elementFormats?.['series:0']).toEqual({ reset: true });
		expect(readChartElementFormat(next, { kind: 'series', seriesIndex: 0 }).reset).toBe(true);
	});
});

describe('format pane options', () => {
	it('sets and clears axis options', () => {
		let next = withChartAxisOptions(bar(), 1, {
			min: 0,
			max: 50,
			crosses: 10,
			reverseOrder: true,
			logBase: 10,
		});
		expect(next.axes?.[1]).toMatchObject({
			min: 0,
			max: 50,
			crossesAt: 10,
			orientation: 'maxMin',
			logScale: true,
			logBase: 10,
		});
		next = withChartAxisOptions(next, 1, { min: null, crosses: 'max', logBase: null });
		expect(next.axes?.[1].min).toBeUndefined();
		expect(next.axes?.[1]).toMatchObject({ crosses: 'max', logScale: false });
		expect(next.axes?.[1].crossesAt).toBeUndefined();
		expect(() => withChartAxisOptions(bar(), 1, { max: Number.NaN })).toThrow(RangeError);
	});

	it('sets legend, data label and series options', () => {
		expect(withChartLegendOptions(bar(), { position: 't', overlay: false }).style).toMatchObject({
			legendPosition: 't',
			legendOverlay: false,
		});
		expect(
			withChartDataLabelOptions(bar(), { showCategory: true, position: 'outEnd' }).style,
		).toMatchObject({
			hasDataLabels: true,
			dataLabels: { showCategory: true, position: 'outEnd' },
		});
		expect(withChartSeriesOptions(bar(), { gapWidth: 80, overlap: -10 })).toMatchObject({
			barGapWidth: 80,
			barOverlap: -10,
		});
		expect(() => withChartSeriesOptions(bar(), { gapWidth: 900 })).toThrow(RangeError);
	});
});
