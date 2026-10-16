// `c:chartSpace` writer: a series' trendlines and error bars in the sequences of ECMA-376 Part 1,
// 21.2: CT_Trendline (name, spPr, trendlineType, order, period, forward, backward, intercept,
// dispRSqr, dispEq, trendlineLbl, extLst), CT_TrendlineLbl (layout, tx, numFmt, spPr, txPr,
// extLst) and CT_ErrBars (errDir, errBarType, errValType, noEndCap, plus, minus, val, spPr, extLst).
// The schema requires a trendline type, a bar type and a value type: a model without one is
// written with the schema's usual default (`linear`, `both`, `fixedVal`).
import type { ChartErrorBars, ChartTrendline, ChartTrendlineLabel } from './model-analysis';
import { layoutXml, numberFormatXml, shapePropertiesXml, textBodyXml } from './write-shape';
import { chartTextXml, dataSourceXml } from './write-text';
import { elementXml, escapeText, raw, valXml, type ChartWriteContext } from './write-util';

function trendlineLabelXml(
	context: ChartWriteContext,
	label: ChartTrendlineLabel | undefined,
): string {
	if (!label) return '';
	return elementXml(
		'trendlineLbl',
		layoutXml(context, label.layout) +
			chartTextXml(context, label.tx, false) +
			numberFormatXml(label.numberFormat) +
			shapePropertiesXml(context, label.spPr) +
			textBodyXml(context, label.txPr) +
			raw(context, label.extLst),
	);
}

/** One trendline (`c:trendline`). */
export function trendlineXml(context: ChartWriteContext, trendline: ChartTrendline): string {
	return elementXml(
		'trendline',
		(trendline.name === undefined ? '' : `<c:name>${escapeText(trendline.name)}</c:name>`) +
			shapePropertiesXml(context, trendline.spPr) +
			valXml('trendlineType', trendline.trendlineType ?? 'linear') +
			valXml('order', trendline.order) +
			valXml('period', trendline.period) +
			valXml('forward', trendline.forward) +
			valXml('backward', trendline.backward) +
			valXml('intercept', trendline.intercept) +
			valXml('dispRSqr', trendline.displayRSquared) +
			valXml('dispEq', trendline.displayEquation) +
			trendlineLabelXml(context, trendline.label) +
			raw(context, trendline.extLst),
	);
}

/** One set of error bars (`c:errBars`). */
export function errorBarsXml(context: ChartWriteContext, bars: ChartErrorBars): string {
	return elementXml(
		'errBars',
		valXml('errDir', bars.direction) +
			valXml('errBarType', bars.barType ?? 'both') +
			valXml('errValType', bars.valueType ?? 'fixedVal') +
			valXml('noEndCap', bars.noEndCap) +
			dataSourceXml('plus', bars.plus) +
			dataSourceXml('minus', bars.minus) +
			valXml('val', bars.value) +
			shapePropertiesXml(context, bars.spPr) +
			raw(context, bars.extLst),
	);
}

/** The trendlines and error bars of a series, in schema order (between `c:dLbls` and `c:cat`). */
export const seriesAnalysisXml = (
	context: ChartWriteContext,
	series: { trendlines?: ChartTrendline[]; errorBars?: ChartErrorBars[] },
): string =>
	(series.trendlines ?? []).map((trendline) => trendlineXml(context, trendline)).join('') +
	(series.errorBars ?? []).map((bars) => errorBarsXml(context, bars)).join('');
