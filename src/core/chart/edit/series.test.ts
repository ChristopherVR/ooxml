import { describe, expect, it } from 'vitest';
import { chartSourceValues } from '../data-cache';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { chartSpaceSeries } from './draft';
import {
	addChartCategory,
	addChartSeries,
	moveChartSeries,
	removeChartCategory,
	removeChartSeries,
	renameChartSeries,
	setChartCategories,
	setChartDataPointValue,
	setChartSeriesValues,
} from './series';
import { columnChart } from './test-support/column-chart';

const values = (space: ReturnType<typeof columnChart>, i: number) =>
	chartSourceValues(chartSpaceSeries(space)[i]!.values);

describe('series edits', () => {
	it('adds a series after the highest idx and order, with the shared categories', () => {
		const next = addChartSeries(columnChart(), { name: 'Forecast', values: [1, null, 3] });
		const added = chartSpaceSeries(next)[2]!;
		expect(added).toMatchObject({ index: 2, order: 2, tx: { value: 'Forecast' } });
		expect(added.categories?.formula).toBe('Sheet1!$A$2:$A$4');
		expect(added.values).toMatchObject({ kind: 'numLit', cache: { pointCount: 3 } });
		expect(values(next, 2)).toEqual([1, '', 3]);
		expect(parseChartSpace(writeChartSpace(next)).issues).toEqual([]);
	});

	it('removes, renames and reorders series', () => {
		expect(chartSpaceSeries(removeChartSeries(columnChart(), 0)).map((s) => s.tx?.text)).toEqual([
			'Actual',
		]);
		const renamed = chartSpaceSeries(renameChartSeries(columnChart(), 1, 'Real'))[1]!;
		expect(renamed.tx).toEqual({ value: 'Real', text: 'Real' });
		const moved = chartSpaceSeries(moveChartSeries(columnChart(), 1, 0));
		expect(moved.map((s) => [s.tx?.text, s.index, s.order])).toEqual([
			['Actual', 1, 0],
			['Plan', 0, 1],
		]);
		expect(() => removeChartSeries(columnChart(), 5)).toThrow(RangeError);
		expect(() => renameChartSeries(columnChart(), '__proto__' as never, 'x')).toThrow(RangeError);
	});

	it('replaces values keeping the reference and format code', () => {
		const next = setChartSeriesValues(columnChart(), 0, [5, 6, 7, 8]);
		const source = chartSpaceSeries(next)[0]!.values!;
		expect(source).toMatchObject({ kind: 'numRef', formula: 'Sheet1!$B$2:$B$4' });
		expect(source.cache).toMatchObject({ formatCode: 'General', pointCount: 4 });
		expect(values(next, 0)).toEqual([5, 6, 7, 8]);
		expect(values(setChartDataPointValue(next, 0, 1, 60), 0)).toEqual([5, 60, 7, 8]);
		expect(() => setChartSeriesValues(columnChart(), 0, [Number.NaN])).toThrow(RangeError);
	});

	it('edits the shared categories of every series', () => {
		const next = setChartCategories(columnChart(), ['A', 'B']);
		for (const series of chartSpaceSeries(next))
			expect(chartSourceValues(series.categories)).toEqual(['A', 'B']);
		const added = addChartCategory(columnChart(), 'West');
		expect(chartSourceValues(chartSpaceSeries(added)[1]!.categories)).toEqual([
			'North',
			'South',
			'East',
			'West',
		]);
		expect(values(added, 1)).toEqual([12, 18, 33, 0]);
		const removed = removeChartCategory(columnChart(), 0);
		expect(chartSourceValues(chartSpaceSeries(removed)[0]!.categories)).toEqual(['South', 'East']);
		expect(values(removed, 0)).toEqual([20, 30]);
	});
});
