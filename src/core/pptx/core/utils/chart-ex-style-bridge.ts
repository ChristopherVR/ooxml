// PowerPoint compatibility styles: resolve theme colours/fonts and retain the raw region-map cache.
// Chart data, axes and family options are read from ChartExSpace by chart-ex-from-neutral.
import type { PptxChartData, PptxChartSeries, XmlObject } from '../types';
import type { ColorParserLike, XmlLookupLike } from './chart-cx-parser';
import { parseCxAxes } from './chart-cx-axis-parser';
import { parseCxDataLabels } from './chart-cx-data-labels';
import { parseCxRegionMapOptions } from './chart-cx-region-map';
import { extractCxSeriesColor } from './chart-cx-series-color';
export function readChartExStyleBridge(
	plotArea: XmlObject,
	lookup: XmlLookupLike,
	space: XmlObject | undefined,
	color?: ColorParserLike,
	font?: (raw: string) => string,
) {
	const region = lookup.getChildByLocalName(plotArea, 'plotAreaRegion');
	const nodes = lookup.getChildrenArrayByLocalName(region, 'series');
	const data = lookup.getChildrenArrayByLocalName(
		lookup.getChildByLocalName(space, 'chartData'),
		'data',
	);
	const series = nodes.map((node, index): PptxChartSeries => {
		const result: PptxChartSeries = { name: `Series ${index + 1}`, values: [0] };
		const paint = extractCxSeriesColor(node, lookup, color);
		if (paint) result.color = paint;
		const labels = parseCxDataLabels(node, lookup, color, font);
		if (labels?.labels.length) result.dataLabels = labels.labels;
		if (labels?.options && Object.keys(labels.options).length)
			result.dataLabelOptions = labels.options;
		const id = lookup.getChildByLocalName(node, 'dataId')?.['@_val'];
		const source =
			lookup.getChildByLocalName(node, 'data') ??
			data.find((data) => String(data['@_id']) === String(id));
		const map = parseCxRegionMapOptions(node, source, lookup, color);
		if (map) result.regionMapOptions = map;
		return result;
	});
	const axes = parseCxAxes(plotArea, lookup, color, font);
	return {
		categories: [] as string[],
		series,
		...(axes ? { chartData: { axes } as Partial<PptxChartData> } : {}),
	};
}
