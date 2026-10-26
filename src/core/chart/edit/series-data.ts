// Data sources of edited series: a reference keeps its formula and gets a fresh cache (as Office
// refreshes the cache from the cells), a source without one becomes a literal. Values are numbers or
// blanks (`null`, written as a missing point, as Office writes an empty cell); categories are text,
// or numbers when every category is a number.
import type { ChartSeriesValue } from '../data-sheet';
import type { ChartDataSource, ChartDataCache } from '../model-series';

export type ChartCategoryValue = string | number | null;

function points(values: readonly (string | number | null)[], previous?: ChartDataCache) {
	const formats = new Map(
		previous?.points.flatMap((point) =>
			point.index !== undefined && point.formatCode !== undefined
				? [[point.index, point.formatCode] as const]
				: [],
		),
	);
	return values.flatMap((value, index) => {
		const formatCode = formats.get(index);
		return value === null || value === ''
			? []
			: [{ index, value: String(value), ...(formatCode === undefined ? {} : { formatCode }) }];
	});
}

/** A number source holding `values`, keeping the reference and format code of `previous`. */
export function numberSource(
	values: readonly ChartSeriesValue[],
	previous: ChartDataSource | undefined,
	formula?: string,
): ChartDataSource {
	for (const value of values)
		if (value !== null && !Number.isFinite(value))
			throw new RangeError('Chart values must be finite numbers or null');
	const reference = formula ?? (previous?.kind === 'numRef' ? previous.formula : undefined);
	return {
		kind: reference !== undefined ? 'numRef' : 'numLit',
		...(reference !== undefined ? { formula: reference } : {}),
		cache: {
			type: 'number',
			formatCode: previous?.cache?.formatCode ?? 'General',
			pointCount: values.length,
			points: points(values, previous?.cache),
		},
	};
}

/** A category source holding `values` (numbers when every value is one), keeping the reference. */
export function categorySource(
	values: readonly ChartCategoryValue[],
	previous: ChartDataSource | undefined,
	formula?: string,
): ChartDataSource {
	const numeric = values.length > 0 && values.every((value) => typeof value === 'number');
	const reference =
		formula ??
		(previous?.kind === 'numRef' ||
		previous?.kind === 'strRef' ||
		previous?.kind === 'multiLvlStrRef'
			? previous.formula
			: undefined);
	const kind = numeric
		? reference !== undefined
			? 'numRef'
			: 'numLit'
		: reference !== undefined
			? 'strRef'
			: 'strLit';
	return {
		kind,
		...(reference !== undefined ? { formula: reference } : {}),
		cache: {
			type: numeric ? 'number' : 'string',
			...(numeric ? { formatCode: previous?.cache?.formatCode ?? 'General' } : {}),
			pointCount: values.length,
			points: points(values, numeric ? previous?.cache : undefined),
		},
	};
}
