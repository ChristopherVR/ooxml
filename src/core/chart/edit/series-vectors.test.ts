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
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { writeChartSpaceOverSource } from '../keep-source';
const data = { categories: ['A', 'B'], series: [{ name: 'Sales', values: [1, 2] }] };
it('retains numeric point formats, opaque series extensions and stable combo series pointers', () => {
	const space = setChartSeriesBubbleSizes(
		newChartSpace(
			{ ...data, series: [...data.series, { name: 'Other', values: [3, 4] }] },
			{ kind: 'bubble' },
		),
		1,
		[5, 6],
	);
	const group = space.plotArea.groups[0]!;
	const selected = group.series.pop()!;
	space.plotArea.groups.push({ ...structuredClone(group), series: [selected] });
	selected.xValues!.cache!.points[0]!.formatCode = '0.000';
	selected.bubbleSizes!.cache!.points[1]!.formatCode = '0%';
	selected.extLst =
		'<c:extLst><c:ext uri="opaque-series"><x:keep xmlns:x="urn:private"/></c:ext></c:extLst>';
	const xml = writeChartSpace(space),
		loaded = parseChartSpace(xml).chartSpace;
	const oldFirst = structuredClone(loaded.plotArea.groups[0]!.series[0]!);
	const changed = setChartSeriesBubbleSizes(
		setChartSeriesXValues(loaded, 1, [10, 20]),
		1,
		[30, 40],
	);
	const saved = writeChartSpaceOverSource(xml, changed),
		reopened = parseChartSpace(saved).chartSpace;
	expect(reopened.plotArea.groups[0]!.series[0]).toEqual(oldFirst);
	expect(reopened.plotArea.groups[1]!.series[0]!.xValues!.cache!.points).toEqual([
		{ index: 0, value: '10', formatCode: '0.000' },
		{ index: 1, value: '20' },
	]);
	expect(reopened.plotArea.groups[1]!.series[0]!.bubbleSizes!.cache!.points[1]).toMatchObject({
		value: '40',
		formatCode: '0%',
	});
	expect(saved).toContain('opaque-series');
	expect(saved).toContain('urn:private');
	expect(reopened.plotArea.groups[1]!.series[0]!.xValues!.formula).toBe(selected.xValues!.formula);
});
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
	expect(
		setChartTypedCategories(space, ['Text', '']).plotArea.groups[0]!.series[0]!.categories,
	).toMatchObject({
		kind: 'strRef',
		formula: source.formula,
		cache: { type: 'string', pointCount: 2, points: [{ index: 0, value: 'Text' }] },
	});
});
it('converts numeric categories to text without renumbering authored series or changing their formulas', () => {
	const space = setChartNumericCategories(
		newChartSpace(
			{ ...data, series: [...data.series, { name: 'Second', values: [3, 4] }] },
			{ kind: 'line' },
		),
		[2011, 2012],
	);
	space.plotArea.groups[0]!.series.forEach((series, index) => {
		series.index = index + 1;
		series.order = index + 1;
	});
	const source = writeChartSpace(space),
		parsed = parseChartSpace(source).chartSpace;
	const changed = setChartTypedCategories(parsed, ['Cat 1', 'Cat 2']);
	const saved = parseChartSpace(writeChartSpaceOverSource(source, changed)).chartSpace;
	expect(saved.plotArea.groups[0]!.series.map((series) => series.index)).toEqual([1, 2]);
	for (const [index, series] of saved.plotArea.groups[0]!.series.entries()) {
		expect(series.categories).toMatchObject({
			kind: 'strRef',
			formula: parsed.plotArea.groups[0]!.series[index]!.categories!.formula,
			cache: {
				type: 'string',
				points: [
					{ index: 0, value: 'Cat 1' },
					{ index: 1, value: 'Cat 2' },
				],
			},
		});
		expect(series.values).toEqual(parsed.plotArea.groups[0]!.series[index]!.values);
	}
	expect(() =>
		setChartTypedCategories(newChartSpace(data, { kind: 'scatter' }), ['Text', 'Other']),
	).toThrow('finite');
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
