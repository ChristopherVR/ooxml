/**
 * Trendlines of PowerPoint's chart Format pane (Format Data Series >
 * Trendlines): the shared pane's neutral `trendline` intent applied to a
 * `PptxChartData`, one trendline at a time, with Add trendline appending one.
 * The type, polynomial order, moving-average period, forecasts and the
 * equation and R-squared switches follow `setChartTrendline` in
 * `ooxml-core/chart`; the line colour, width, dash and label layout PowerPoint
 * keeps on a trendline are left as they are.
 *
 * @module sdk/chart-trendline-operations
 */
import {
	chartTrendlineKind,
	chartTrendlineType,
	type ChartTrendline,
	type ChartTrendlineEdit,
} from '../../../../chart/index';
import type { PptxChartData, PptxChartTrendline } from '../../types';
import { chartSupportsErrorBars } from './chart-error-bar-operations';

/** The trendlines of series `seriesIndex` as the shared pane reads them. */
export function chartSeriesTrendlinesView(
	data: PptxChartData,
	seriesIndex: number,
): ChartTrendline[] {
	return (data.series[seriesIndex]?.trendlines ?? []).map((trend) => {
		const view: ChartTrendline = { trendlineType: chartTrendlineType(trend.trendlineType) };
		if (trend.name !== undefined) view.name = trend.name;
		for (const key of ['order', 'period', 'forward', 'backward', 'intercept'] as const)
			if (trend[key] !== undefined) view[key] = trend[key];
		if (trend.displayEq !== undefined) view.displayEquation = trend.displayEq;
		if (trend.displayRSq !== undefined) view.displayRSquared = trend.displayRSq;
		return view;
	});
}

const integer = (value: number, min: number, max: number) =>
	Number.isInteger(value) && value >= min && value <= max;

/**
 * `data` with trendline `index` of series `seriesIndex` changed by `edit`
 * (index = the trendline count adds one, linear unless `edit` names a type),
 * or removed for `null`. Throws a `RangeError` for a family without
 * trendlines, a missing series or trendline, or an invalid option.
 */
export function withChartSeriesTrendline(
	data: PptxChartData,
	seriesIndex: number,
	index: number,
	edit: ChartTrendlineEdit | null,
): PptxChartData {
	const series = data.series[seriesIndex];
	if (!series) throw new RangeError(`No series at index ${seriesIndex}`);
	if (!chartSupportsErrorBars(data))
		throw new RangeError('This chart family does not support trendlines');
	const list = [...(series.trendlines ?? [])];
	if (
		!Number.isSafeInteger(index) ||
		index < 0 ||
		index > list.length ||
		(edit === null && index === list.length)
	)
		throw new RangeError('No trendline at this index');
	if (edit === null) list.splice(index, 1);
	else {
		const next: PptxChartTrendline = { ...(list[index] ?? { trendlineType: 'linear' }) };
		if (edit.type !== undefined) {
			const kind = chartTrendlineKind(edit.type);
			if (!kind) throw new RangeError('Unsupported trendline type');
			next.trendlineType = kind;
			if (kind !== 'polynomial') delete next.order;
			if (kind !== 'movingAvg') delete next.period;
		}
		for (const key of ['order', 'period', 'forward', 'backward', 'intercept'] as const) {
			const value = edit[key];
			if (value === undefined) continue;
			if (value === null) {
				delete next[key];
				continue;
			}
			if (
				!Number.isFinite(value) ||
				(key === 'order' && !integer(value, 2, 6)) ||
				(key === 'period' && !integer(value, 2, 2147483647)) ||
				((key === 'forward' || key === 'backward') && value < 0)
			)
				throw new RangeError(`Invalid trendline ${key}`);
			next[key] = value;
		}
		if (next.trendlineType === 'polynomial') next.order ??= 2;
		else if (next.order !== undefined)
			throw new RangeError('Only polynomial trendlines have an order');
		if (next.trendlineType === 'movingAvg') next.period ??= 2;
		else if (next.period !== undefined) throw new RangeError('Only moving averages have a period');
		if (edit.name !== undefined) {
			if (edit.name === null) delete next.name;
			else next.name = edit.name;
		}
		for (const [from, to] of [
			['displayEquation', 'displayEq'],
			['displayRSquared', 'displayRSq'],
		] as const) {
			const value = edit[from];
			if (value === null) delete next[to];
			else if (value !== undefined) next[to] = value;
		}
		list[index] = next;
	}
	return {
		...data,
		series: data.series.map((item, i) =>
			i === seriesIndex ? { ...item, trendlines: list } : item,
		),
	};
}
