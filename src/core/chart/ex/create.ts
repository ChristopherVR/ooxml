// Create literal ChartEx data through the same neutral operations used for loaded charts.
import type { ChartDataSheet } from '../data-sheet';
import type { ChartExDataSheet } from '../edit/chartex-sheet';
import type { ChartExGeography, ChartExSpace } from './model';
import { setChartExTitle, setChartExLegend, setChartExBinning } from '../edit/chartex';
import { setChartExCategories, setChartExSeriesValues } from '../edit/chartex-data';
import { setChartExLayout } from '../edit/chartex-structure';
import { newChartExGeography, toChartExPareto, type ChartExFamily } from '../edit/chartex-family';

export interface CreateChartExOptions {
	title?: string | undefined;
	legend?: { visible: boolean; position?: string | undefined } | undefined;
	/**
	 * Region map: the culture and, optionally, projection and area. The part gets no provider
	 * cache (`cx:geoCache`); see {@link newChartExGeography}.
	 */
	geography?: Partial<Omit<ChartExGeography, 'geoCacheXml' | 'extLst'>>;
}

/**
 * A new extended chart of `family` over literal data. Every family but box and whisker and
 * histogram takes one series. A region map's categories are place names, resolved by the
 * consumer; a pareto chart aggregates its categories (see `toChartExPareto`).
 */
export function createChartExSpace(
	data: ChartDataSheet | ChartExDataSheet,
	family: ChartExFamily,
	options: CreateChartExOptions = {},
): ChartExSpace {
	if (!data.series.length) throw new RangeError('A chart needs at least one series.');
	if (data.series.length > 1 && !['boxWhisker', 'clusteredColumn'].includes(family))
		throw new RangeError('This chart family requires a single series.');
	const map = family === 'regionMap';
	let space: ChartExSpace = {
		chartData: { data: [] },
		plotArea: {
			axes: [],
			series: data.series.map((series, index) => ({
				layoutId: map ? 'regionMap' : 'funnel',
				tx: { value: series.name },
				dataId: index,
				dataPoints: [],
				axisIds: [],
				...(map ? { layoutProperties: { geography: newChartExGeography(options.geography) } } : {}),
			})),
		},
	};
	if (!map) space = setChartExLayout(space, family === 'pareto' ? 'clusteredColumn' : family);
	if (family === 'clusteredColumn')
		for (let index = 0; index < data.series.length; index++)
			space = setChartExBinning(space, index, {});
	space = setChartExCategories(
		space,
		'categoryLevels' in data && data.categoryLevels.length ? data.categoryLevels : data.categories,
	);
	data.series.forEach((series, index) => {
		space = setChartExSeriesValues(space, index, series.values);
	});
	if (family === 'pareto') space = toChartExPareto(space);

	if (options.title !== undefined) space = setChartExTitle(space, options.title);
	if (options.legend) space = setChartExLegend(space, options.legend);
	return space;
}
