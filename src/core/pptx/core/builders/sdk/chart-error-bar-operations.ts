/**
 * Error bars of PowerPoint's chart Format pane (Format Data Series > Error
 * Bars): the shared pane's neutral `errorBars` intent applied to a
 * `PptxChartData`. The amount (fixed value, percentage, standard deviation,
 * standard error or custom), the direction, the end caps and the custom
 * positive and negative values. PowerPoint keeps custom amounts as literal
 * values here (`c:numLit`), so a custom source is always a list of numbers; a
 * series' x error bars and its line formatting are left as they are.
 *
 * @module sdk/chart-error-bar-operations
 */
import type { ChartDataSource, ChartErrorBars, ChartErrorBarsEdit } from '../../../../chart/index';
import type { PptxChartData, PptxChartErrBars, PptxChartSeries } from '../../types';

/** Chart families PowerPoint draws Y error bars for in the Format pane (2-D, not stacked). */
const ERROR_BAR_TYPES = new Set(['bar', 'line', 'area', 'scatter', 'bubble']);

/** Whether series of `data` take error bar edits from the Format pane. */
export function chartSupportsErrorBars(data: PptxChartData): boolean {
	return (
		ERROR_BAR_TYPES.has(data.chartType) &&
		data.grouping !== 'stacked' &&
		data.grouping !== 'percentStacked'
	);
}

const literal = (values: readonly number[]): ChartDataSource => ({
	kind: 'numLit',
	cache: {
		type: 'number',
		formatCode: 'General',
		pointCount: values.length,
		points: values.map((value, index) => ({ index, value: String(value) })),
	},
});

const numbers = (source: ChartDataSource): number[] => {
	const values = (source.cache?.points ?? []).map((point) => Number(point.value));
	if (!values.length || !values.every((value) => Number.isFinite(value) && value >= 0))
		throw new RangeError('Custom error values must be nonnegative numbers');
	return values;
};

const yBars = (series: PptxChartSeries | undefined) =>
	(series?.errBars ?? []).filter((bars) => (bars.direction ?? 'y') === 'y');

/** A series' Y error bars as the shared pane reads them, or undefined without (or with several). */
export function chartSeriesErrorBarsView(
	data: PptxChartData,
	seriesIndex: number,
): { errorBars?: ChartErrorBars; multiple: boolean } {
	const bars = yBars(data.series[seriesIndex]);
	const one = bars.length === 1 ? bars[0] : undefined;
	if (!one) return { multiple: bars.length > 1 };
	const view: ChartErrorBars = {
		direction: 'y',
		barType: one.barType,
		valueType: one.valType,
		...(one.noEndCap === undefined ? {} : { noEndCap: one.noEndCap }),
		...(one.val === undefined ? {} : { value: one.val }),
		...(one.customPlus?.length ? { plus: literal(one.customPlus) } : {}),
		...(one.customMinus?.length ? { minus: literal(one.customMinus) } : {}),
	};
	return { errorBars: view, multiple: false };
}

/**
 * `data` with the Y error bars of series `seriesIndex` changed by `edit`
 * (created as a fixed value of 1, both sides, when missing), or removed for
 * `null`. Throws a `RangeError` for a family without error bars, several Y
 * error bar sets, a custom source that is not nonnegative numbers, or an
 * unknown series.
 */
export function withChartSeriesErrorBars(
	data: PptxChartData,
	seriesIndex: number,
	edit: ChartErrorBarsEdit | null,
): PptxChartData {
	const series = data.series[seriesIndex];
	if (!series) throw new RangeError(`No series at index ${seriesIndex}`);
	if (!chartSupportsErrorBars(data))
		throw new RangeError('This chart family does not support error bars');
	const list = series.errBars ?? [];
	const at = list.flatMap((bars, index) => ((bars.direction ?? 'y') === 'y' ? [index] : []));
	if (at.length > 1) throw new RangeError('Ambiguous error-bar direction');
	let errBars = list.filter((_, index) => index !== at[0]);
	if (edit !== null) {
		const next: PptxChartErrBars = {
			...(at.length
				? list[at[0]!]!
				: { direction: 'y', barType: 'both', valType: 'fixedVal', val: 1 }),
		};
		if (edit.barType !== undefined) next.barType = edit.barType;
		if (edit.valueType !== undefined) next.valType = edit.valueType;
		if (edit.noEndCap === null) delete next.noEndCap;
		else if (edit.noEndCap !== undefined) next.noEndCap = edit.noEndCap;
		if (edit.value === null) delete next.val;
		else if (edit.value !== undefined) {
			if (!Number.isFinite(edit.value) || edit.value < 0)
				throw new RangeError('The error amount must be a nonnegative number');
			next.val = edit.value;
		}
		if (edit.plus === null) delete next.customPlus;
		else if (edit.plus) next.customPlus = numbers(edit.plus);
		if (edit.minus === null) delete next.customMinus;
		else if (edit.minus) next.customMinus = numbers(edit.minus);
		if (next.valType !== 'cust') {
			delete next.customPlus;
			delete next.customMinus;
		} else delete next.val;
		errBars = at.length
			? list.map((bars, index) => (index === at[0] ? next : bars))
			: [...list, next];
	}
	return {
		...data,
		series: data.series.map((item, index) => (index === seriesIndex ? { ...item, errBars } : item)),
	};
}
