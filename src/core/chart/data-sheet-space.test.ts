import { describe, expect, it } from 'vitest';
import {
	applyChartDataSheetEditToSpace,
	applyChartDataSheetEditsToSpace,
	chartSpaceDataSheet,
} from './data-sheet-space';
import { columnChart } from './edit/test-support/column-chart';

describe('chart data sheet over a ChartSpace', () => {
	it('reads names, categories and values from the part caches', () => {
		expect(chartSpaceDataSheet(columnChart())).toEqual({
			categories: ['North', 'South', 'East'],
			series: [
				{ name: 'Plan', values: [10, 20, 30] },
				{ name: 'Actual', values: [12, 18, 33] },
			],
		});
	});

	it('applies edits through the series operations, keeping series formatting', () => {
		const space = applyChartDataSheetEditsToSpace(columnChart(), [
			{ kind: 'remove-category', category: 1 },
			{ kind: 'category-label', category: 0, label: 'N' },
			{ kind: 'remove-series', series: 0 },
			{ kind: 'value', series: 0, category: 1, value: 7 },
		]);
		expect(chartSpaceDataSheet(space)).toEqual({
			categories: ['N', 'East'],
			series: [{ name: 'Actual', values: [12, 7] }],
		});
		expect(space.plotArea.groups[0]!.series[0]!.spPr?.fill).toMatchObject({ kind: 'solid' });
	});

	it('leaves the chart unchanged for an edit the sheet refuses', () => {
		const space = columnChart();
		expect(applyChartDataSheetEditToSpace(space, { kind: 'remove-series', series: 5 })).toBe(space);
		const one = applyChartDataSheetEditToSpace(space, { kind: 'remove-series', series: 0 });
		expect(applyChartDataSheetEditToSpace(one, { kind: 'remove-series', series: 0 })).toBe(one);
	});
});
