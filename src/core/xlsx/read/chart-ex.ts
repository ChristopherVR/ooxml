// Extended charts (waterfall, histogram...) on a sheet. Excel wraps the graphic frame in
// `mc:AlternateContent` with a "not available" fallback shape and writes the data as hidden
// `_xlchart.v1.N` defined names. The frame stays an unsupported drawing object (its anchor XML is
// kept verbatim on save); the families the chart painter draws also carry their summary, so the
// grid can paint them from live cells.
import { parseChartExSpace } from '../../chart/ex/parse';
import { chartExSummary } from '../../chart/render/summary-from-chartex';
import { NS, children, first, relAttr, type XmlElement } from '../../xml/index';
import type { DrawingAnchor, UnsupportedDrawingObject } from '../model';
import type { SourceIndex } from './package';

/** The `xdr:graphicFrame` inside the `mc:Choice` of an alternate-content anchor. */
function choiceFrame(content: XmlElement): XmlElement | undefined {
	for (const choice of children(content, 'Choice', NS.mc)) {
		const frame = first(choice, 'graphicFrame', NS.xdr);
		if (frame) return frame;
	}
	return undefined;
}

/**
 * Reads an extended chart frame, or undefined when the content is not one. `describe` is what the
 * object is called when the painter does not draw its family.
 */
export function readChartExFrame(
	source: SourceIndex,
	partName: string,
	content: XmlElement,
	anchor: DrawingAnchor,
	sourceXml: string,
	name: string | undefined,
): UnsupportedDrawingObject | undefined {
	const frame = choiceFrame(content);
	const data = first(first(frame, 'graphic', NS.a), 'graphicData', NS.a);
	if (!data?.getAttribute('uri')?.includes('chartex')) return undefined;
	const chart = first(data, 'chart', NS.cx);
	const rel = source.rels(partName).get(relAttr(chart, 'id') ?? '');
	const target = rel ? source.target(partName, rel) : undefined;
	const xml = target ? source.text(target) : undefined;
	const object: UnsupportedDrawingObject = {
		kind: 'unsupported',
		anchor,
		description: 'chart (newer type)',
		sourceXml,
	};
	if (!xml) return object;
	try {
		const parsed = parseChartExSpace(xml);
		object.chartExSpace = parsed.chartExSpace;
		object.chartExIssues = parsed.issues;
		if (target) object.chartExPartName = target;
		const extended = chartExSummary(xml);
		if (extended?.layoutId) object.description = `chart (${extended.layoutId})`;
		if (extended?.summary) {
			object.chartEx = {
				...extended.summary,
				kind: 'chart',
				anchor,
				series: extended.summary.series,
			};
			if (target) object.chartEx.partName = target;
			if (name) object.chartEx.name = name;
		}
	} catch {
		// A part that cannot be read stays a labelled frame.
	}
	return object;
}
