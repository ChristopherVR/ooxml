// Region map series options: Format Data Series > Series Options in Office (map projection, map
// area, map labels). Values are the chartex schema's enumerations ([MS-ODRAWXML]
// ST_GeoProjectionType, ST_GeoMappingLevel, ST_RegionLabelLayout); `auto` removes the attribute
// so Office chooses. The provider cache is never touched.
import type { ChartExSpace } from '../ex/model';
import { chartExSeriesAt, editChartExSpace } from './chartex';

export const CHART_EX_MAP_PROJECTIONS = ['mercator', 'miller', 'robinson', 'albers'] as const;
export const CHART_EX_MAP_AREAS = [
	'dataOnly',
	'postalCode',
	'county',
	'state',
	'countryRegion',
	'countryRegionList',
	'world',
] as const;
export const CHART_EX_MAP_LABELS = ['none', 'bestFitOnly', 'showAll'] as const;

export interface ChartExRegionMapOptions {
	projectionType?: (typeof CHART_EX_MAP_PROJECTIONS)[number] | 'auto';
	viewedRegionType?: (typeof CHART_EX_MAP_AREAS)[number] | 'auto';
	regionLabelLayout?: (typeof CHART_EX_MAP_LABELS)[number] | 'auto';
}

function check(value: string | undefined, allowed: readonly string[], what: string): void {
	if (value !== undefined && value !== 'auto' && !allowed.includes(value))
		throw new RangeError(`Unsupported map ${what}.`);
}

/** Sets the projection, area and label layout of the region map series at `index`. */
export function setChartExRegionMap(
	space: ChartExSpace,
	index: number,
	options: ChartExRegionMapOptions,
): ChartExSpace {
	if (chartExSeriesAt(space, index).layoutId !== 'regionMap')
		throw new RangeError('The series is not a region map.');
	check(options.projectionType, CHART_EX_MAP_PROJECTIONS, 'projection');
	check(options.viewedRegionType, CHART_EX_MAP_AREAS, 'area');
	check(options.regionLabelLayout, CHART_EX_MAP_LABELS, 'label layout');
	return editChartExSpace(space, (draft) => {
		const layout = (chartExSeriesAt(draft, index).layoutProperties ??= {});
		if (options.regionLabelLayout === 'auto') delete layout.regionLabelLayout;
		else if (options.regionLabelLayout) layout.regionLabelLayout = options.regionLabelLayout;
		for (const key of ['projectionType', 'viewedRegionType'] as const) {
			const value = options[key];
			if (value === undefined) continue;
			if (!layout.geography)
				throw new Error('The region map has no geography to change; Office authors it.');
			if (value === 'auto') delete layout.geography[key];
			else layout.geography[key] = value;
		}
	});
}
