// Create literal ChartEx data through the same neutral operations used for loaded charts.
import type { ChartDataSheet } from '../data-sheet';
import type { ChartExDataSheet } from '../edit/chartex-sheet';
import type { ChartExSpace } from './model';
import { setChartExTitle, setChartExLegend, setChartExBinning } from '../edit/chartex';
import { setChartExCategories, setChartExSeriesValues } from '../edit/chartex-data';
import { setChartExLayout, type ChartExLayout } from '../edit/chartex-structure';

/** Geographic charts require producer-authored geography metadata and cannot be created here. */
export function createChartExSpace(
	data: ChartDataSheet | ChartExDataSheet,
	layout: Exclude<ChartExLayout, 'regionMap'>,
	options: { title?: string; legend?: { visible: boolean; position?: string } } = {},
): ChartExSpace {
	if (!data.series.length) throw new RangeError('A chart needs at least one series.');
	if (data.series.length > 1 && !['boxWhisker', 'clusteredColumn'].includes(layout))
		throw new RangeError('This chart family requires a single series.');
	let space: ChartExSpace = {
		chartData: { data: [] },
		plotArea: {
			axes: [],
			series: data.series.map((series, index) => ({
				layoutId: 'funnel',
				tx: { value: series.name },
				dataId: index,
				dataPoints: [],
				axisIds: [],
			})),
		},
	};
	space = setChartExLayout(space, layout);
	if (layout === 'clusteredColumn')
		for (let index = 0; index < data.series.length; index++)
			space = setChartExBinning(space, index, {});
	space = setChartExCategories(
		space,
		'categoryLevels' in data && data.categoryLevels.length ? data.categoryLevels : data.categories,
	);
	data.series.forEach((series, index) => {
		space = setChartExSeriesValues(space, index, series.values);
	});

	if (options.title !== undefined) space = setChartExTitle(space, options.title);
	if (options.legend) space = setChartExLegend(space, options.legend);
	return space;
}
