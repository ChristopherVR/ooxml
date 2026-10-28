// Builds the spreadsheet chart object from the shared chart summary (`chartSummaryFromSpace` over
// `chart/parseChartSpace`), adding the anchor and each series' SpreadsheetML colour. The parsed
// `ChartSpace` is kept on the object: edits change it and the summary is derived from it again
// (`chartObjectFromSpace`). Direct formatting metadata is read from the part's XML by
// `chart/readChartFormatting`.
import { parseXml, type XmlElement } from '../../xml/index';
import type { ChartObject, ChartSeries, Color, DrawingAnchor } from '../model';
import { readChartFormatting } from '../../chart/read-formatting';
import { parseChartSpace } from '../../chart/parse-space';
import { writeChartSpace } from '../../chart/write-space';
import { chartIssuesLostOnWrite } from '../../chart/keep-source';
import type { ChartSpace } from '../../chart/model';
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

/** What a chart object keeps beside its part: placement, package part and session state. */
export type ChartObjectFrame = Pick<ChartObject, 'anchor'> &
	Partial<
		Pick<
			ChartObject,
			'partName' | 'name' | 'quickLayout' | 'colorPalette' | 'styleDefinition' | 'editBlocked'
		>
	>;

/**
 * The chart object showing `space`: the painter's summary, SpreadsheetML series colours and the
 * direct formatting, read from `root` (the part as written, `writeChartSpace(space)` when absent),
 * with the frame's placement and session fields.
 */
export function chartObjectFromSpace(
	space: ChartSpace,
	frame: ChartObjectFrame,
	root?: XmlElement,
): ChartObject {
	const summary = chartSummaryFromSpace(space);
	const series = summary.series.map((source): ChartSeries => {
		const color = legacyColor(source.drawingColor);
		return color ? { ...source, color } : source;
	});
	const object: ChartObject = {
		kind: 'chart',
		anchor: frame.anchor,
		chartType: summary.chartType,
		series,
		showLegend: summary.showLegend,
	};
	if (frame.partName !== undefined) object.partName = frame.partName;
	if (summary.date1904 !== undefined) object.date1904 = summary.date1904;
	if (summary.plotLayout) object.plotLayout = summary.plotLayout;
	if (summary.axes) object.axes = summary.axes;
	if (summary.axisIds) object.axisIds = summary.axisIds;
	if (summary.grouping) object.grouping = summary.grouping;
	if (summary.barGapWidth !== undefined) object.barGapWidth = summary.barGapWidth;
	if (summary.barOverlap !== undefined) object.barOverlap = summary.barOverlap;
	if (summary.title !== undefined) object.title = summary.title;
	if (summary.legendPosition) object.legendPosition = summary.legendPosition;
	if (summary.extraGroups) object.extraGroups = summary.extraGroups;
	if (summary.bubble) object.bubble = summary.bubble;
	if (summary.stock) object.stock = summary.stock;
	if (frame.name) object.name = frame.name;
	if (frame.quickLayout !== undefined) object.quickLayout = frame.quickLayout;
	if (frame.colorPalette !== undefined) object.colorPalette = frame.colorPalette;
	if (frame.styleDefinition) object.styleDefinition = frame.styleDefinition;
	if (frame.editBlocked) object.editBlocked = frame.editBlocked;
	const formatting = readChartFormatting(
		root ?? parseXml(writeChartSpace(space), { label: 'XLSX chart' }).documentElement,
	);
	if (formatting) object.formatting = formatting;
	object.chartSpace = space;
	return object;
}

/** Reads the modelled summary of a `c:chartSpace` part. */
export function parseChart(
	xml: string,
	anchor: DrawingAnchor,
	partName: string,
	name?: string,
): ChartObject {
	const root = parseXml(xml, { label: 'XLSX chart' }).documentElement;
	const { chartSpace, issues } = parseChartSpace(root);
	const lost = chartIssuesLostOnWrite(issues);
	const editBlocked = lost.length
		? `Chart ${partName} has content the editor does not model (${[
				...new Set(lost.map((issue) => issue.element ?? issue.code)),
			].join(', ')}); it is shown and saved unchanged but cannot be edited.`
		: undefined;
	return chartObjectFromSpace(
		chartSpace,
		{ anchor, partName, ...(name ? { name } : {}), ...(editBlocked ? { editBlocked } : {}) },
		root,
	);
}
