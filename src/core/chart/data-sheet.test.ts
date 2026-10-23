import { describe, expect, it } from 'vitest';
import { parseLocaleNumber } from '../text/locale-number';
import { parseTsv } from '../text/tsv';
import {
	applyChartDataSheetEdit,
	applyChartDataSheetEdits,
	chartDataSheetCellEdit,
	chartDataSheetPasteEdits,
	type ChartDataSheet,
} from './data-sheet';

const sheet: ChartDataSheet = {
	categories: ['Q1', 'Q2'],
	series: [
		{ name: 'North', values: [1, 2] },
		{ name: 'South', values: [3] },
	],
};
const en = (raw: string) => parseLocaleNumber(raw, 'en-US');

describe('chart data sheet edits', () => {
	it('writes values, padding a ragged series', () => {
		const next = applyChartDataSheetEdit(sheet, {
			kind: 'value',
			series: 1,
			category: 1,
			value: 9,
		});
		expect(next?.series[1]?.values).toEqual([3, 9]);
		expect(sheet.series[1]?.values).toEqual([3]);
	});

	it('names new rows and columns the way PowerPoint does', () => {
		const next = applyChartDataSheetEdits(sheet, [
			{ kind: 'add-series' },
			{ kind: 'add-category' },
		]);
		expect(next.categories).toEqual(['Q1', 'Q2', 'Cat 3']);
		expect(next.series.map((entry) => entry.name)).toEqual(['North', 'South', 'Series 3']);
		expect(next.series[1]?.values).toEqual([3, null, 0]);
	});

	it('never removes the last series or category, and refuses out-of-range edits', () => {
		const one: ChartDataSheet = { categories: ['A'], series: [{ name: 'S', values: [1] }] };
		expect(applyChartDataSheetEdit(one, { kind: 'remove-series', series: 0 })).toBeNull();
		expect(applyChartDataSheetEdit(one, { kind: 'remove-category', category: 0 })).toBeNull();
		expect(
			applyChartDataSheetEdit(sheet, { kind: 'series-name', series: 5, name: 'x' }),
		).toBeNull();
		expect(
			applyChartDataSheetEdit(sheet, { kind: 'value', series: 0, category: 0, value: Number.NaN }),
		).toBeNull();
		const removed = applyChartDataSheetEdit(sheet, { kind: 'remove-category', category: 0 });
		expect(removed).toEqual({
			categories: ['Q2'],
			series: [
				{ name: 'North', values: [2] },
				{ name: 'South', values: [null] },
			],
		});
	});

	it('maps a grid cell to the edit it means', () => {
		expect(chartDataSheetCellEdit({ row: 0, col: 2 }, 'West', en)).toEqual({
			kind: 'series-name',
			series: 1,
			name: 'West',
		});
		expect(chartDataSheetCellEdit({ row: 2, col: 0 }, 'H2', en)).toEqual({
			kind: 'category-label',
			category: 1,
			label: 'H2',
		});
		expect(chartDataSheetCellEdit({ row: 1, col: 1 }, '1,234.5', en)).toEqual({
			kind: 'value',
			series: 0,
			category: 0,
			value: 1234.5,
		});
		expect(chartDataSheetCellEdit({ row: 1, col: 1 }, 'abc', en)).toBeNull();
		expect(chartDataSheetCellEdit({ row: 0, col: 0 }, 'x', en)).toBeNull();
	});

	it('pastes a block, growing the sheet and skipping cells that are not numbers', () => {
		const block = parseTsv('East\tWest\r\n5\tx\r\n7\t8\r\n');
		const edits = chartDataSheetPasteEdits(sheet, { row: 0, col: 2 }, block, en);
		expect(edits.slice(0, 1)).toEqual([{ kind: 'add-series' }]);
		const next = applyChartDataSheetEdits(sheet, edits);
		expect(next.series.map((entry) => entry.name)).toEqual(['North', 'East', 'West']);
		expect(next.series[1]?.values).toEqual([5, 7]);
		expect(next.series[2]?.values).toEqual([0, 8]);
		expect(chartDataSheetPasteEdits(sheet, { row: 1, col: 1 }, [], en)).toEqual([]);
	});
});
