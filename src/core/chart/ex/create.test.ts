import { expect, it } from 'vitest';
import { createChartExSpace } from './create';
import { parseChartExSpace } from './parse';
import { writeChartExSpace } from './write';
import { chartExSpaceDataSheet } from '../edit/chartex-sheet';
import { setChartExBinning } from '../edit/chartex';
import { setChartExSeriesValues } from '../edit/chartex-data';
import { setChartExLayout } from '../edit/chartex-structure';
const data = {
	categories: ['A', 'B'],
	categoryLevels: [
		['A', 'B'],
		['Parent', 'Parent'],
	],
	series: [{ name: 'Sales', values: [1, null] }],
};
it.each(['waterfall', 'clusteredColumn', 'boxWhisker', 'treemap', 'sunburst', 'funnel'] as const)(
	'creates %s with reusable hierarchy and blank caches',
	(layout) => {
		const space = createChartExSpace(data, layout, { title: 'Created', legend: { visible: true } });
		const result = parseChartExSpace(writeChartExSpace(space));
		expect(result.issues).toEqual([]);
		expect(result.chartExSpace.plotArea.series[0]!.layoutId).toBe(layout);
		expect(chartExSpaceDataSheet(result.chartExSpace)).toMatchObject(data);
		expect(space.plotArea.axes.length).toBe(
			['waterfall', 'clusteredColumn', 'boxWhisker'].includes(layout) ? 2 : 0,
		);
	},
);
it('rejects invalid bin values before changing the source', () => {
	const space = createChartExSpace(data, 'clusteredColumn');
	for (const binning of [
		{ binSize: NaN },
		{ binSize: 0 },
		{ binCount: 1.5 },
		{ overflow: Infinity },
	])
		expect(() => setChartExBinning(space, 0, binning)).toThrow();
	expect(space.plotArea.series[0]!.layoutProperties).toEqual({ binning: {} });
});
it('keeps additional numeric dimensions when editing the displayed value level', () => {
	const space = createChartExSpace(data, 'boxWhisker');
	const dimension = space.chartData.data[0]!.dimensions.find((d) => d.kind === 'numDim')!;
	dimension.levels.push({
		pointCount: 2,
		points: [{ index: 0, value: '99' }],
		name: 'Extra sample',
	});
	const changed = setChartExSeriesValues(space, 0, [3, 4]);
	expect(changed.chartData.data[0]!.dimensions.find((d) => d.kind === 'numDim')!.levels[1]).toEqual(
		dimension.levels[1],
	);
});
it('switches to a histogram with numeric binning rather than an extended column chart', () => {
	const space = setChartExLayout(createChartExSpace(data, 'waterfall'), 'clusteredColumn');
	expect(space.plotArea.series[0]!.layoutProperties).toEqual({ binning: {} });
	expect(writeChartExSpace(space)).toContain('<cx:binning');
});
