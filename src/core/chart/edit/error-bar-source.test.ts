import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { chartSummaryFromSpace } from '../render/summary-from-space';
import { setChartErrorBars, setChartTrendline } from './analysis';
import {
	chartErrorBarSourceFromText,
	chartErrorBarSourceText,
	chartSpaceRangeNumbers,
} from './error-bar-source';
import { applyChartEditIntent } from './intent';
import { columnChart } from './test-support/column-chart';

const reopen = (space: ReturnType<typeof columnChart>) => {
	const result = parseChartSpace(writeChartSpace(space));
	expect(result.issues).toEqual([]);
	return result.chartSpace;
};

describe('custom error-bar sources', () => {
	it('reads literal lists, repeating a single value over every point', () => {
		const one = chartErrorBarSourceFromText('={2}', { pointCount: 3 });
		expect(one?.kind).toBe('numLit');
		expect(one?.cache?.points.map((p) => p.value)).toEqual(['2', '2', '2']);
		const list = chartErrorBarSourceFromText('1; 2.5, 3', { pointCount: 3 });
		expect(list?.cache?.points.map((p) => p.value)).toEqual(['1', '2.5', '3']);
		expect(chartErrorBarSourceText(list)).toBe('1, 2.5, 3');
		for (const bad of ['', '1,,2', '-1', 'abc', '1e999'])
			expect(chartErrorBarSourceFromText(bad, { pointCount: 3 })).toBeUndefined();
	});

	it('resolves ranges only through a resolver, from the cells the chart caches by default', () => {
		const space = columnChart();
		const resolve = chartSpaceRangeNumbers(space);
		expect(resolve('Sheet1!$C$2:$C$4')).toEqual([12, 18, 33]);
		expect(resolve("'Sheet1'!C3")).toEqual([18]);
		// Category text and uncached cells are unknown.
		expect(resolve('Sheet1!$A$2:$A$4')).toBeUndefined();
		expect(resolve('Sheet1!$D$2:$D$4')).toBeUndefined();
		expect(resolve('Sheet1!$B$2:$C$4')).toBeUndefined();
		expect(chartErrorBarSourceFromText('Sheet1!$C$2:$C$4', { pointCount: 3 })).toBeUndefined();
		const source = chartErrorBarSourceFromText('=Sheet1!$C$2:$C$4', {
			pointCount: 3,
			resolveRange: resolve,
		});
		expect(source).toMatchObject({ kind: 'numRef', formula: 'Sheet1!$C$2:$C$4' });
		expect(chartErrorBarSourceText(source)).toBe('Sheet1!$C$2:$C$4');
	});

	it('writes custom range and literal error bars that reopen and paint per point', () => {
		const space = columnChart();
		const plus = chartErrorBarSourceFromText('Sheet1!$C$2:$C$4', {
			pointCount: 3,
			resolveRange: chartSpaceRangeNumbers(space),
		})!;
		const minus = chartErrorBarSourceFromText('1, 2, 3', { pointCount: 3 })!;
		const edited = applyChartEditIntent(space, {
			op: 'errorBars',
			seriesIndex: 0,
			direction: 'y',
			edit: { valueType: 'cust', plus, minus },
		});
		const xml = writeChartSpace(edited);
		expect(xml).toContain(
			'<c:plus><c:numRef><c:f>Sheet1!$C$2:$C$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>',
		);
		expect(xml).toContain(
			'<c:minus><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="3"/>',
		);
		const bars = reopen(edited).plotArea.groups[0]!.series[0]!.errorBars![0]!;
		expect(bars).toMatchObject({ valueType: 'cust', plus: { formula: 'Sheet1!$C$2:$C$4' } });
		expect(bars.minus?.cache?.points.map((p) => p.value)).toEqual(['1', '2', '3']);
		const summary = chartSummaryFromSpace(reopen(edited)).series[0]!;
		expect(JSON.stringify(summary)).toContain('"customPlus":[12,18,33]');
		// Switching away from custom drops both sources.
		const fixed = setChartErrorBars(edited, 0, 'y', { valueType: 'fixedVal' });
		expect(writeChartSpace(fixed)).not.toContain('<c:plus>');
	});
});

describe('several trendlines on one series', () => {
	it('adds, edits and removes trendlines by index and writes each in order', () => {
		let space = setChartTrendline(columnChart(), 0, 0, { type: 'linear' });
		space = setChartTrendline(space, 0, 1, { type: 'movingAvg', period: 2 });
		space = setChartTrendline(space, 0, 2, { type: 'poly', order: 3, displayEquation: true });
		space = setChartTrendline(space, 0, 1, { period: 3 });
		const types = (s: typeof space) =>
			reopen(s).plotArea.groups[0]!.series[0]!.trendlines!.map((t) => [
				t.trendlineType,
				t.period ?? t.order,
			]);
		expect(types(space)).toEqual([
			['linear', undefined],
			['movingAvg', 3],
			['poly', 3],
		]);
		space = setChartTrendline(space, 0, 0, null);
		expect(types(space)).toEqual([
			['movingAvg', 3],
			['poly', 3],
		]);
		expect(() => setChartTrendline(space, 0, 3, { type: 'linear' })).toThrow(RangeError);
	});
});
