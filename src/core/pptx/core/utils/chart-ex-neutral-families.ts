// PowerPoint's family-specific view fields from the shared ChartEx model.
import type { ChartExSeries } from '../../../chart';
import type { PptxChartSeries } from '../types';
export function chartExFamilyFields(series: ChartExSeries): Partial<PptxChartSeries> {
	const layout = series.layoutProperties;
	const out: Partial<PptxChartSeries> = {};
	if (series.layoutId === 'clusteredColumn' || series.layoutId === 'paretoLine') {
		const bin = layout?.binning;
		out.histogramOptions = {
			layout: series.layoutId === 'paretoLine' ? 'pareto' : 'histogram',
			...(bin?.binSize === undefined ? {} : { binSize: bin.binSize }),
			...(bin?.binCount !== undefined && Number.isSafeInteger(bin.binCount) && bin.binCount >= 0
				? { binCount: bin.binCount }
				: {}),
			...(['l', 'r'].includes(bin?.intervalClosed ?? '')
				? { intervalClosed: bin!.intervalClosed as 'l' | 'r' }
				: {}),
			...(bin?.underflow === undefined ? {} : { underflow: bin.underflow }),
			...(bin?.overflow === undefined ? {} : { overflow: bin.overflow }),
			...(layout?.aggregation ? { aggregateByCategory: true } : {}),
		};
	}
	if (
		series.layoutId === 'waterfall' &&
		layout &&
		(layout.visibility?.connectorLines !== undefined || layout.subtotals !== undefined)
	)
		out.waterfallOptions = {
			...(layout.visibility?.connectorLines === undefined
				? {}
				: { connectorLines: layout.visibility.connectorLines }),
			...(layout.subtotals === undefined
				? {}
				: { subtotalIndices: layout.subtotals.filter((i) => Number.isInteger(i) && i >= 0) }),
		};
	if (series.layoutId === 'boxWhisker') {
		const method = layout?.statistics?.quartileMethod;
		const flags = layout?.visibility;
		const options: PptxChartSeries['boxWhiskerOptions'] = {
			quartileMethod: method === 'inclusive' || method === 'exclusive' ? method : undefined,
			showMeanLine: flags?.meanLine,
			showMeanMarker: flags?.meanMarker,
			showInnerPoints: flags?.nonoutliers,
			showOutlierPoints: flags?.outliers,
		};
		if (Object.values(options).some((value) => value !== undefined))
			out.boxWhiskerOptions = options;
	}
	const parent = layout?.parentLabelLayout?.trim();
	if (parent === 'none' || parent === 'banner' || parent === 'overlapping')
		out.treemapOptions = { parentLabelLayout: parent };
	return out;
}
