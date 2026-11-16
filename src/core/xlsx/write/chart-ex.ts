// The parts of an extended chart inserted in this session: the `chartExN.xml` part (its data read
// from sheet ranges, so no embedded workbook), the chart style and colour style parts Office
// requires beside it, and the drawing frame. The frame sits in `mc:AlternateContent` under
// `mc:Choice Requires="cx1"`, as Excel writes it; no fallback shape is written.
import { CHART_EX_PART_CONTENT_TYPE, CHART_EX_RELATIONSHIP_TYPE } from '../../chart/ex/new-part';
import {
	CHART_STYLE_CONTENT_TYPE,
	CHART_STYLE_REL,
	chartExColorStyleXml,
	chartExStyleXml,
} from '../../chart/ex/style-part';
import { CHART_COLOR_STYLE_CONTENT_TYPE, CHART_COLOR_STYLE_REL } from '../../chart/color-style';
import type { ChartExSpace } from '../../chart/ex/model';
import { writeChartExSpace } from '../../chart/ex/write';
import { NS } from '../../xml/index';
import { RelationshipSet, type PackageWriter } from './package-writer';

const CX1 = 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/** Writes the chart's parts and returns the frame for the drawing anchor. */
export function writeInsertedChartEx(
	writer: PackageWriter,
	space: ChartExSpace,
	drawingRels: RelationshipSet,
	target: (part: string) => string,
	id: number,
): string {
	const reserved = writer.source ? new Set(writer.source.parts.keys()) : undefined;
	const part = writer.uniqueName((n) => `xl/charts/chartEx${n}.xml`, reserved);
	writer.add(part, writeChartExSpace(space), CHART_EX_PART_CONTENT_TYPE);
	const style = writer.uniqueName((n) => `xl/charts/style${n}.xml`, reserved);
	writer.add(style, chartExStyleXml(), CHART_STYLE_CONTENT_TYPE);
	const colors = writer.uniqueName((n) => `xl/charts/colors${n}.xml`, reserved);
	writer.add(colors, chartExColorStyleXml(), CHART_COLOR_STYLE_CONTENT_TYPE);
	const rels = new RelationshipSet();
	rels.add(CHART_STYLE_REL, style.slice('xl/charts/'.length));
	rels.add(CHART_COLOR_STYLE_REL, colors.slice('xl/charts/'.length));
	writer.rels(part, rels);
	const relId = drawingRels.add(CHART_EX_RELATIONSHIP_TYPE, target(part));
	return (
		`<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:cx1="${CX1}" Requires="cx1">` +
		`<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
		'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
		`<a:graphic><a:graphicData uri="${NS.cx}"><cx:chart xmlns:cx="${NS.cx}" r:id="${relId}"/></a:graphicData></a:graphic></xdr:graphicFrame>` +
		'</mc:Choice></mc:AlternateContent>'
	);
}
