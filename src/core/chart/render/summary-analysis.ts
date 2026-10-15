// The trendlines and error bars of a summary series, from the neutral model's `c:trendline` and
// `c:errBars`: what the painter needs to fit, size and stroke them. Elements whose type cannot be
// read are skipped, as Office skips them.
import { chartSourceValues } from '../data-cache';
import { chartTrendlineKind, type ChartErrorBars, type ChartTrendline } from '../model-analysis';
import type { ChartDataSource } from '../model-series';
import type { ChartSummaryErrorBars, ChartSummaryTrendline } from './summary';

function summaryTrendline(source: ChartTrendline): ChartSummaryTrendline | undefined {
	const kind = chartTrendlineKind(source.trendlineType);
	if (!kind) return undefined;
	const out: ChartSummaryTrendline = { kind };
	if (source.order !== undefined) out.order = source.order;
	if (source.period !== undefined) out.period = source.period;
	if (source.forward !== undefined) out.forward = source.forward;
	if (source.backward !== undefined) out.backward = source.backward;
	if (source.intercept !== undefined) out.intercept = source.intercept;
	if (source.name !== undefined) out.name = source.name;
	if (source.spPr?.line) out.line = source.spPr.line;
	return out;
}

/** Custom amounts by point index; blanks and text are 0. */
const amounts = (source: ChartDataSource | undefined): number[] =>
	chartSourceValues(source).map((value) => (typeof value === 'number' ? value : 0));

function summaryErrorBars(source: ChartErrorBars): ChartSummaryErrorBars | undefined {
	if (!source.valueType) return undefined;
	const out: ChartSummaryErrorBars = {
		direction: source.direction ?? 'y',
		barType: source.barType ?? 'both',
		valueType: source.valueType,
	};
	if (source.value !== undefined) out.value = source.value;
	if (source.valueType === 'cust') {
		out.customPlus = amounts(source.plus);
		out.customMinus = amounts(source.minus);
	}
	if (source.noEndCap !== undefined) out.noEndCap = source.noEndCap;
	if (source.spPr?.line) out.line = source.spPr.line;
	return out;
}

/** The summary's trendlines and error bars of one series, each absent when there are none. */
export function summaryAnalysis(series: {
	trendlines?: ChartTrendline[];
	errorBars?: ChartErrorBars[];
}): { trendlines?: ChartSummaryTrendline[]; errorBars?: ChartSummaryErrorBars[] } {
	const trendlines = (series.trendlines ?? [])
		.map(summaryTrendline)
		.filter((entry): entry is ChartSummaryTrendline => entry !== undefined);
	const errorBars = (series.errorBars ?? [])
		.map(summaryErrorBars)
		.filter((entry): entry is ChartSummaryErrorBars => entry !== undefined);
	return {
		...(trendlines.length ? { trendlines } : {}),
		...(errorBars.length ? { errorBars } : {}),
	};
}
