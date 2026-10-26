// Typed category and XY/bubble vectors retain references until the host allocates new workbook ranges.
import type { ChartSpace } from '../model';
import { chartCacheValues } from '../data-cache';
import type { ChartSeriesValue } from '../data-sheet';
import { editChartSpace, chartSeriesEntryAt, chartSeriesEntries } from './draft';
import { numberSource, categorySource } from './series-data';

export function setChartSeriesXValues(
	space: ChartSpace,
	index: number,
	values: readonly ChartSeriesValue[],
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const { group, series } = chartSeriesEntryAt(draft, index);
		if (group.kind !== 'scatter' && group.kind !== 'bubble')
			throw new RangeError('Only scatter and bubble charts have X values.');
		series.xValues = numberSource(values, series.xValues, series.xValues?.formula);
	});
}
export function setChartSeriesBubbleSizes(
	space: ChartSpace,
	index: number,
	values: readonly ChartSeriesValue[],
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const { group, series } = chartSeriesEntryAt(draft, index);
		if (group.kind !== 'bubble') throw new RangeError('Only bubble charts have size values.');
		series.bubbleSizes = numberSource(values, series.bubbleSizes, series.bubbleSizes?.formula);
	});
}
/** Leaves parent hierarchy levels authored when only the displayed (innermost) labels change. */
export function setChartTypedCategories(space: ChartSpace, values: readonly string[]): ChartSpace {
	return editChartSpace(space, (draft) => {
		for (const { group, series } of chartSeriesEntries(draft)) {
			const key = group.kind === 'scatter' || group.kind === 'bubble' ? 'xValues' : 'categories';
			const source = series[key];
			if (source?.levels) {
				if (
					Math.max(
						source.levelPointCount ?? 0,
						...source.levels.map((level) => chartCacheValues(level).length),
					) !== values.length
				)
					throw new RangeError('Growing a hierarchy requires every category level.');
				const first = source.levels[0];
				if (!first) throw new RangeError('The category hierarchy has no levels.');
				first.pointCount = values.length;
				first.points = values.flatMap((value, index) => (value === '' ? [] : [{ index, value }]));
			} else if (source?.cache?.type === 'number' || key === 'xValues') {
				const numbers = values.map((value) => (value.trim() === '' ? null : Number(value)));
				series[key] = numberSource(numbers, source, source?.formula);
			} else series[key] = categorySource(values, source);
		}
	});
}
/** Complete hierarchy, leaf level first. Unequal levels cannot map to one rectangular worksheet range. */
export function setChartCategoryLevels(
	space: ChartSpace,
	levels: readonly (readonly string[])[],
): ChartSpace {
	if (!levels.length || levels.some((level) => level.length !== levels[0]!.length))
		throw new RangeError('Category levels must have matching lengths.');
	return editChartSpace(space, (draft) => {
		for (const { group, series } of chartSeriesEntries(draft)) {
			if (group.kind === 'scatter' || group.kind === 'bubble')
				throw new RangeError('XY charts do not support category hierarchies.');
			const previous = series.categories;
			series.categories = {
				kind: 'multiLvlStrRef',
				...(previous?.formula ? { formula: previous.formula } : {}),
				levelPointCount: levels[0]!.length,
				levels: levels.map((values, index) => ({
					...previous?.levels?.[index],
					type: 'string',
					pointCount: values.length,
					points: values.flatMap((value, index) => (value === '' ? [] : [{ index, value }])),
				})),
			};
		}
	});
}
/** Numeric/date categories keep number-cache formats and blanks. Date-system flags stay unchanged. */
export function setChartNumericCategories(
	space: ChartSpace,
	values: readonly ChartSeriesValue[],
	formatCode?: string,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		for (const { group, series } of chartSeriesEntries(draft)) {
			if (group.kind === 'scatter' || group.kind === 'bubble')
				throw new RangeError('Use per-series X values for XY charts.');
			series.categories = numberSource(values, series.categories, series.categories?.formula);
			if (formatCode !== undefined) series.categories.cache!.formatCode = formatCode;
		}
	});
}
