import { expect, it } from 'vitest';
import { newChartSpace } from './new-chart';
import {
	setChartSeriesXValues,
	setChartSeriesBubbleSizes,
	setChartCategoryLevels,
	setChartTypedCategories,
	setChartNumericCategories,
} from './series-vectors';
import { rebaseChartData } from '../workbook/rebase-data';
const data = { categories: ['A', 'B'], series: [{ name: 'Sales', values: [1, 2] }] };
it('keeps numeric categories, date formats, sparse blanks and source references', () => {
	const space = setChartNumericCategories(
		newChartSpace(data, { kind: 'bar' }),
		[45000, null],
		'yyyy-mm-dd',
	);
	const next = setChartTypedCategories(space, ['45001', '']);
	const source = next.plotArea.groups[0]!.series[0]!.categories!;
	expect(source).toMatchObject({
		kind: 'numRef',
		formula: 'Sheet1!$A$2:$A$3',
		cache: { formatCode: 'yyyy-mm-dd', pointCount: 2, points: [{ index: 0, value: '45001' }] },
	});
	expect(() => setChartTypedCategories(space, ['Invalid', ''])).toThrow();
});
it('keeps hierarchy parents, accepts a full growth rectangle and rebases it outermost first', () => {
	const space = setChartCategoryLevels(newChartSpace(data, { kind: 'bar' }), [
		['A', 'B'],
		['Parent', 'Parent'],
	]);
	const next = setChartTypedCategories(space, ['Leaf', 'B']);
	expect(next.plotArea.groups[0]!.series[0]!.categories!.levels![1]!.points[0]!.value).toBe(
		'Parent',
	);
	expect(() => setChartTypedCategories(next, ['A', 'B', 'C'])).toThrow('every category level');
	const grown = setChartCategoryLevels(next, [
		['Leaf', 'B', 'C'],
		['Parent', 'Parent', 'Other'],
	]);
	const result = rebaseChartData(grown, 'SafeNewSheet');
	expect(result.rows[1]!.slice(0, 2)).toEqual(['Parent', 'Leaf']);
	expect(result.chartSpace.plotArea.groups[0]!.series[0]!.categories!.formula).toContain(
		'$A$2:$B$4',
	);
});
it('edits independent XY and bubble-size vectors with validation and new-sheet allocation', () => {
	const space = newChartSpace(data, { kind: 'bubble' });
	const x = setChartSeriesXValues(space, 0, [10, 20]);
	const bubble = setChartSeriesBubbleSizes(x, 0, [3, null]);
	const result = rebaseChartData(bubble, 'NewData');
	expect(result.rows[1]).toEqual([10, 1, 3]);
	expect(result.rows[2]).toEqual([20, 2, null]);
	expect(() => setChartSeriesXValues(space, 0, [Infinity])).toThrow();
	expect(() => setChartSeriesBubbleSizes(newChartSpace(data, { kind: 'bar' }), 0, [1])).toThrow(
		'Only bubble',
	);
});
