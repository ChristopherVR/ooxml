import { describe, expect, it } from 'vitest';

import type { PptxChartData, XmlObject } from '../../types';
import { applySeriesErrBarsToXml } from '../../utils/chart-errbars-serializer';
import { chartSeriesErrorBarsView, withChartSeriesErrorBars } from './chart-error-bar-operations';

const data = (): PptxChartData => ({
	chartType: 'bar',
	categories: ['A', 'B'],
	series: [{ name: 'S', values: [1, 2] }],
	grouping: 'clustered',
});

const literal = (values: number[]) => ({
	kind: 'numLit' as const,
	cache: {
		type: 'number' as const,
		pointCount: values.length,
		points: values.map((value, index) => ({ index, value: String(value) })),
	},
});

describe('withChartSeriesErrorBars', () => {
	it('writes custom error values the save writes as c:numLit and reads back', () => {
		const next = withChartSeriesErrorBars(data(), 0, {
			valueType: 'cust',
			plus: literal([2, 3]),
			minus: literal([1, 1]),
		});
		expect(next.series[0]!.errBars).toEqual([
			{ direction: 'y', barType: 'both', valType: 'cust', customPlus: [2, 3], customMinus: [1, 1] },
		]);
		const node: XmlObject = { 'c:idx': { '@_val': '0' }, 'c:val': {} };
		applySeriesErrBarsToXml(node, next.series[0]!.errBars!, (key) => key.split(':').pop()!);
		const bars = node['c:errBars'] as XmlObject;
		expect(bars['c:errValType']).toEqual({ '@_val': 'cust' });
		expect(bars['c:plus']).toEqual({
			'c:numLit': {
				'c:ptCount': { '@_val': '2' },
				'c:pt': [
					{ '@_idx': '0', 'c:v': '2' },
					{ '@_idx': '1', 'c:v': '3' },
				],
			},
		});
		expect(chartSeriesErrorBarsView(next, 0).errorBars).toMatchObject({
			valueType: 'cust',
			plus: { cache: { points: [{ value: '2' }, { value: '3' }] } },
		});
	});

	it('refuses families without error bars and bad amounts', () => {
		expect(() => withChartSeriesErrorBars({ ...data(), chartType: 'pie' }, 0, {})).toThrow(
			RangeError,
		);
		expect(() => withChartSeriesErrorBars(data(), 0, { value: -1 })).toThrow(RangeError);
		expect(() => withChartSeriesErrorBars(data(), 3, {})).toThrow(RangeError);
	});
});
