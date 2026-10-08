// The spreadsheet entry to the shared chart view: live cell references and SpreadsheetML series
// colours on top of the neutral `chart/render` view model.
import {
	chartSummaryView,
	type ChartViewModel,
	type EvaluateRef,
} from '../../chart/render/chart-view';
import type { ChartObject, Workbook } from '../model';
import { resolveColor } from './colors';

export {
	categoryTotals,
	type ChartRefValue,
	type ChartSeriesView,
	type ChartViewModel,
	type EvaluateRef,
	type ValueAxisView,
} from '../../chart/render/chart-view';

/**
 * Resolves a chart's series from live cells (`evaluateRef` returns the values of a reference such
 * as `Sheet1!$B$2:$B$9`, flattened row-major), falling back to the values cached in the chart part.
 */
export function chartView(
	workbook: Workbook,
	sheetIndex: number,
	chart: ChartObject,
	evaluateRef: EvaluateRef,
): ChartViewModel {
	void sheetIndex;
	const theme = workbook.theme;
	return chartSummaryView(chart, theme, {
		evaluateRef,
		seriesColor: (series) => resolveColor(series.color, theme),
	});
}
