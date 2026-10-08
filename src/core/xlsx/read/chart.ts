// Builds the spreadsheet chart object from the shared chart summary (`chartSummaryFromSpace` over
// `chart/parseChartSpace`), adding the anchor and each series' SpreadsheetML colour. Direct
// formatting metadata is still read from the part by `chart/readChartFormatting`.
import { parseXml } from '../../xml/index';
import type { ChartObject, ChartSeries, Color, DrawingAnchor } from '../model';
import { readChartFormatting } from '../../chart/read-formatting';
import { parseChartSpace } from '../../chart/parse-space';
import { chartSummaryFromSpace } from '../../chart/render/summary-from-space';
import type { DrawingColor } from '../../drawingml/types';

/** DrawingML scheme colour names to SpreadsheetML theme indices. */
const SCHEME_INDEX: Record<string, number> = {
	bg1: 0,
	lt1: 0,
	tx1: 1,
	dk1: 1,
	bg2: 2,
	lt2: 2,
	tx2: 3,
	dk2: 3,
	accent1: 4,
	accent2: 5,
	accent3: 6,
	accent4: 7,
	accent5: 8,
	accent6: 9,
	hlink: 10,
	folHlink: 11,
};

/** The SpreadsheetML colour of a plain solid fill colour (RGB or a scheme slot), if it has one. */
function legacyColor(color: DrawingColor | undefined): Color | undefined {
	if (color?.kind === 'srgb') return color.value ? { rgb: color.value.toUpperCase() } : undefined;
	if (color?.kind === 'scheme') {
		const index = SCHEME_INDEX[color.value];
		return index === undefined ? undefined : { theme: index };
	}
	return undefined;
}

/** Reads the modelled summary of a `c:chartSpace` part. */
export function parseChart(
	xml: string,
	anchor: DrawingAnchor,
	partName: string,
	name?: string,
): ChartObject {
	const root = parseXml(xml, { label: 'XLSX chart' }).documentElement;
	const { chartSpace } = parseChartSpace(root);
	const summary = chartSummaryFromSpace(chartSpace);
	const series = summary.series.map((source): ChartSeries => {
		const color = legacyColor(source.drawingColor);
		return color ? { ...source, color } : source;
	});
	const object: ChartObject = {
		kind: 'chart',
		anchor,
		chartType: summary.chartType,
		series,
		showLegend: summary.showLegend,
		partName,
	};
	if (summary.grouping) object.grouping = summary.grouping;
	if (summary.barGapWidth !== undefined) object.barGapWidth = summary.barGapWidth;
	if (summary.barOverlap !== undefined) object.barOverlap = summary.barOverlap;
	if (summary.title !== undefined) object.title = summary.title;
	if (summary.legendPosition) object.legendPosition = summary.legendPosition;
	if (name) object.name = name;
	const formatting = readChartFormatting(root);
	if (formatting) object.formatting = formatting;
	return object;
}
