/**
 * chart-data-helpers.test.ts: Vitest unit tests for chart-data-helpers.ts.
 *
 * All tests are pure (no TestBed, no Angular imports).
 *
 * @module angular-viewer/chart-data-helpers.test
 */

import type { ChartPptxElement, PptxChartData } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import {
	patchChartData,
	patchChartStyle,
	setAxisLogScale,
	setAxisTitleStyle,
	setDataPointExplosion,
	setDataPointFill,
	setDataPointMarker,
	setGridlineStyle,
	setSeriesChartType,
	setSeriesColor,
	setSeriesMarker,
} from './chart-data-helpers';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeChart(
	seriesNames: string[],
	categories: string[],
	valuesFn?: (si: number) => number[],
): ChartPptxElement {
	const chartData: PptxChartData = {
		chartType: 'bar',
		categories: [...categories],
		series: seriesNames.map((name, si) => ({
			name,
			values: valuesFn ? valuesFn(si) : categories.map(() => 0),
		})),
	};
	return {
		type: 'chart',
		id: 'ch-1',
		x: 0,
		y: 0,
		width: 400,
		height: 300,
		chartData,
	};
}

// ---------------------------------------------------------------------------
// addSeries
// ---------------------------------------------------------------------------

describe('patchChartStyle', () => {
	it('merges style fields', () => {
		const el = makeChart(['A'], ['x']);
		const result = patchChartStyle(el, { hasLegend: true, legendPosition: 'b' });
		expect(result.chartData?.style?.hasLegend).toBeTruthy();
		expect(result.chartData?.style?.legendPosition).toBe('b');
	});

	it('preserves existing style fields not in patch', () => {
		const el: ChartPptxElement = {
			...makeChart(['A'], ['x']),
			chartData: {
				...makeChart(['A'], ['x']).chartData!,
				style: { hasLegend: true, hasTitle: true },
			},
		};
		const result = patchChartStyle(el, { hasLegend: false });
		expect(result.chartData?.style?.hasTitle).toBeTruthy();
	});
});

// ---------------------------------------------------------------------------
// patchChartData
// ---------------------------------------------------------------------------

describe('patchChartData', () => {
	it('merges arbitrary chartData fields', () => {
		const el = makeChart(['A'], ['x']);
		const result = patchChartData(el, { title: 'My Chart' });
		expect(result.chartData?.title).toBe('My Chart');
	});

	it('uses chartDataChangeType when chartType changes', () => {
		const el = makeChart(['A'], ['x']);
		// Bar -> pie: grouping should be cleared
		const result = patchChartData(el, { chartType: 'pie' });
		expect(result.chartData?.chartType).toBe('pie');
		// pie does not support grouping, so grouping should be undefined/null
		expect(result.chartData?.grouping).toBeUndefined();
	});

	it('does not mutate the original', () => {
		const el = makeChart(['A'], ['x']);
		patchChartData(el, { title: 'Changed' });
		expect(el.chartData?.title).toBeUndefined();
	});
});

// ---------------------------------------------------------------------------
// setSeriesColor
// ---------------------------------------------------------------------------

describe('setSeriesColor', () => {
	it('sets a series colour from a #-prefixed hex', () => {
		const el = makeChart(['A', 'B'], ['x']);
		const result = setSeriesColor(el, 0, '#FF0000');
		expect(result.chartData?.series[0].color).toBe('#FF0000');
	});

	it('normalises a bare hex by prefixing #', () => {
		const el = makeChart(['A'], ['x']);
		const result = setSeriesColor(el, 0, '00FF00');
		expect(result.chartData?.series[0].color).toBe('#00FF00');
	});

	it('clears the colour when passed null', () => {
		const el = makeChart(['A'], ['x']);
		const colored = setSeriesColor(el, 0, '#123456');
		const cleared = setSeriesColor(colored, 0, null);
		expect(cleared.chartData?.series[0].color).toBeUndefined();
	});

	it('only affects the targeted series', () => {
		const el = makeChart(['A', 'B'], ['x']);
		const result = setSeriesColor(el, 1, '#ABCDEF');
		expect(result.chartData?.series[0].color).toBeUndefined();
		expect(result.chartData?.series[1].color).toBe('#ABCDEF');
	});

	it('does not mutate the original', () => {
		const el = makeChart(['A'], ['x']);
		setSeriesColor(el, 0, '#FF0000');
		expect(el.chartData?.series[0].color).toBeUndefined();
	});

	it('returns the element unchanged when chartData is missing', () => {
		const el = { type: 'chart', id: 'x', x: 0, y: 0, width: 1, height: 1 } as ChartPptxElement;
		expect(setSeriesColor(el, 0, '#FFFFFF')).toBe(el);
	});
});

// ---------------------------------------------------------------------------
// Advanced formatting wrappers
// ---------------------------------------------------------------------------

describe('advanced chart formatting wrappers', () => {
	function pieChart(): ChartPptxElement {
		const el = makeChart(['Rev', 'Cost'], ['Q1', 'Q2', 'Q3']);
		el.chartData!.chartType = 'pie';
		return el;
	}

	it('setAxisLogScale enables log scaling immutably', () => {
		const el = makeChart(['Rev'], ['Q1']);
		const result = setAxisLogScale(el, 'valAx', { enabled: true, base: 10 });
		expect(result).not.toBe(el);
		const axis = result.chartData?.axes?.find((a) => a.axisType === 'valAx');
		expect(axis?.logScale).toBeTruthy();
		expect(axis?.logBase).toBe(10);
		expect(el.chartData?.axes).toBeUndefined();
	});

	it('setAxisTitleStyle sets font props', () => {
		const el = makeChart(['Rev'], ['Q1']);
		const result = setAxisTitleStyle(el, 'valAx', {
			fontFamily: 'Arial',
			fontSize: 11,
			fontBold: true,
		});
		const axis = result.chartData?.axes?.find((a) => a.axisType === 'valAx');
		expect(axis?.fontFamily).toBe('Arial');
		expect(axis?.fontSize).toBe(11);
		expect(axis?.fontBold).toBeTruthy();
	});

	it('setGridlineStyle sets gridline line styling', () => {
		const el = makeChart(['Rev'], ['Q1']);
		const result = setGridlineStyle(el, 'valAx', 'major', { color: '#CCCCCC', width: 0.75 });
		const axis = result.chartData?.axes?.find((a) => a.axisType === 'valAx');
		expect(axis?.majorGridlinesSpPr?.strokeColor).toBe('#CCCCCC');
		expect(axis?.majorGridlinesSpPr?.strokeWidth).toBe(0.75);
	});

	it('setSeriesMarker sets and clears a marker', () => {
		const el = makeChart(['Rev'], ['Q1']);
		const set = setSeriesMarker(el, 0, { symbol: 'circle', size: 6, fillColor: '#FF0000' });
		expect(set.chartData?.series[0].marker?.symbol).toBe('circle');
		expect(set.chartData?.series[0].marker?.spPr?.fillColor).toBe('#FF0000');
		const cleared = setSeriesMarker(set, 0, null);
		expect(cleared.chartData?.series[0].marker).toBeUndefined();
	});

	it('setSeriesChartType promotes to combo', () => {
		const el = makeChart(['Rev', 'Cost'], ['Q1']);
		const result = setSeriesChartType(el, 1, 'line');
		expect(result.chartData?.series[1].seriesChartType).toBe('line');
		expect(result.chartData?.chartType).toBe('combo');
	});

	it('setDataPointFill sets and clears a per-point fill', () => {
		const el = pieChart();
		const set = setDataPointFill(el, 0, 1, '#123456');
		expect(set.chartData?.series[0].dataPoints?.find((p) => p.idx === 1)?.spPr?.fillColor).toBe(
			'#123456',
		);
		const cleared = setDataPointFill(set, 0, 1, null);
		// An empty list (not `undefined`) tells the save path the overrides were cleared.
		expect(cleared.chartData?.series[0].dataPoints).toStrictEqual([]);
	});

	it('setDataPointMarker seeds and merges a per-point marker override', () => {
		const el = makeChart(['S'], ['A', 'B', 'C']);

		const seeded = setDataPointMarker(el, 0, 1, { symbol: 'circle' });
		expect(seeded.chartData?.series[0].dataPoints?.find((p) => p.idx === 1)?.marker?.symbol).toBe(
			'circle',
		);
		// The source element must not be touched (immutable contract).
		expect(el.chartData?.series[0].dataPoints).toBeUndefined();

		// A later size edit merges onto the existing override rather than replacing it.
		const sized = setDataPointMarker(seeded, 0, 1, { size: 12 });
		const marker = sized.chartData?.series[0].dataPoints?.find((p) => p.idx === 1)?.marker;
		expect(marker?.symbol).toBe('circle');
		expect(marker?.size).toBe(12);

		const filled = setDataPointMarker(sized, 0, 1, { fillColor: '#ff0000' });
		expect(
			filled.chartData?.series[0].dataPoints?.find((p) => p.idx === 1)?.marker?.spPr?.fillColor,
		).toBe('#ff0000');
	});

	it('setDataPointMarker drops the whole c:dPt when the override is cleared', () => {
		const el = setDataPointMarker(makeChart(['S'], ['A', 'B']), 0, 1, { symbol: 'star' });

		const cleared = setDataPointMarker(el, 0, 1, null);

		// An empty list (not `undefined`) tells the save path the overrides were cleared.
		expect(cleared.chartData?.series[0].dataPoints).toStrictEqual([]);
	});

	it('setDataPointMarker leaves a point fill in place when only the marker is cleared', () => {
		const withFill = setDataPointFill(makeChart(['S'], ['A', 'B']), 0, 1, '#123456');
		const withMarker = setDataPointMarker(withFill, 0, 1, { symbol: 'star' });

		const cleared = setDataPointMarker(withMarker, 0, 1, null);

		const point = cleared.chartData?.series[0].dataPoints?.find((p) => p.idx === 1);
		expect(point?.marker).toBeUndefined();
		expect(point?.spPr?.fillColor).toBe('#123456');
	});

	it('setDataPointExplosion sets a pie slice explosion', () => {
		const el = pieChart();
		const result = setDataPointExplosion(el, 0, 2, 25);
		expect(result.chartData?.series[0].dataPoints?.find((p) => p.idx === 2)?.explosion).toBe(25);
	});

	it('returns the element unchanged when chartData is missing', () => {
		const el: ChartPptxElement = {
			type: 'chart',
			id: 'x',
			x: 0,
			y: 0,
			width: 1,
			height: 1,
			chartData: undefined,
		};
		expect(setAxisLogScale(el, 'valAx', { enabled: true })).toBe(el);
	});
});
