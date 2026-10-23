import { describe, expect, it } from 'vitest';

import type { PptxChartData } from '../types/chart';
import { chartDataApplySheetEdits, chartDataToSheet } from './chart-data-sheet';

const data = {
	chartType: 'bar',
	categories: ['Q1', 'Q2', 'Q3'],
	series: [
		{ name: 'North', values: [1, 2, 3], color: '#FF0000' },
		{ name: 'South', values: [4] },
	],
} as PptxChartData;

describe('chart data sheet adapters', () => {
	it('projects categories, names and values, a ragged series as blanks', () => {
		expect(chartDataToSheet(data)).toEqual({
			categories: ['Q1', 'Q2', 'Q3'],
			series: [
				{ name: 'North', values: [1, 2, 3] },
				{ name: 'South', values: [4, null, null] },
			],
		});
	});

	it('applies edits through the grid policy and keeps series formatting', () => {
		const next = chartDataApplySheetEdits(data, [
			{ kind: 'series-name', series: 0, name: 'East' },
			{ kind: 'value', series: 1, category: 2, value: 7 },
			{ kind: 'category-label', category: 0, label: 'Jan' },
			{ kind: 'add-series' },
		]);
		expect(next?.series[0]).toMatchObject({ name: 'East', color: '#FF0000' });
		expect(next?.series[1]?.values).toEqual([4, 0, 7]);
		expect(next?.categories[0]).toBe('Jan');
		expect(next?.series[2]).toMatchObject({ name: 'Series 3', values: [0, 0, 0] });
	});

	it('returns null when every edit is refused', () => {
		const one = { ...data, series: [data.series[0]] } as PptxChartData;
		expect(chartDataApplySheetEdits(one, [{ kind: 'remove-series', series: 0 }])).toBeNull();
		expect(
			chartDataApplySheetEdits(data, [{ kind: 'value', series: 0, category: 9, value: 1 }]),
		).toBeNull();
	});
});
