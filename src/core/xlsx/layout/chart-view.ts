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
 * A treemap or sunburst keeps its category range, one column per hierarchy level, behind a hidden
 * `_xlchart.v1.N` defined name. The view needs the column count, so the name is replaced by the
 * range it stands for (the name is left alone when it is not a plain reference).
 */
function withHierarchyRanges(workbook: Workbook, chart: ChartObject): ChartObject {
	if (chart.chartType !== 'treemap' && chart.chartType !== 'sunburst') return chart;
	const series = chart.series.map((s) => {
		const ref = s.categoriesRef;
		if (!ref || !/^_xlchart\./i.test(ref)) return s;
		const named = workbook.definedNames.find(
			(n) => n.localSheet === undefined && n.name.toLowerCase() === ref.toLowerCase(),
		);
		return named ? { ...s, categoriesRef: named.formula.replace(/^=/, '') } : s;
	});
	return { ...chart, series };
}

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
	return chartSummaryView(withHierarchyRanges(workbook, chart), theme, {
		evaluateRef,
		seriesColor: (series) => resolveColor(series.color, theme),
	});
}
