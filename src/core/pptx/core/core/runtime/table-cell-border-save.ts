/**
 * Write a table cell's four edge lines (`a:lnL` / `a:lnR` / `a:lnT` /
 * `a:lnB`) from its {@link PptxTableCellStyle}: width, colour and dash. A
 * preset dash is `a:prstDash`, a custom one `a:custDash` with its `a:ds`
 * segments, through the same writer shape outlines use.
 *
 * @module runtime/table-cell-border-save
 */

import type {
	PptxCustomDashSegment,
	PptxTableCellStyle,
	StrokeDashType,
	XmlObject,
} from '../../types';
import { applyDrawingLineDash } from '../../utils/drawing-line-dash';
import { LN_CHILD_ORDER, reorderObjectKeys } from '../../utils/xml-reorder';

interface CellEdge {
	xmlKey: 'a:lnL' | 'a:lnR' | 'a:lnT' | 'a:lnB';
	width: number | undefined;
	color: string | undefined;
	dash: string | undefined;
	segments: PptxCustomDashSegment[] | undefined;
}

function cellEdges(style: PptxTableCellStyle): CellEdge[] {
	return [
		{
			xmlKey: 'a:lnT',
			width: style.borderTopWidth,
			color: style.borderTopColor,
			dash: style.borderTopDash,
			segments: style.borderTopDashSegments,
		},
		{
			xmlKey: 'a:lnB',
			width: style.borderBottomWidth,
			color: style.borderBottomColor,
			dash: style.borderBottomDash,
			segments: style.borderBottomDashSegments,
		},
		{
			xmlKey: 'a:lnL',
			width: style.borderLeftWidth,
			color: style.borderLeftColor,
			dash: style.borderLeftDash,
			segments: style.borderLeftDashSegments,
		},
		{
			xmlKey: 'a:lnR',
			width: style.borderRightWidth,
			color: style.borderRightColor,
			dash: style.borderRightDash,
			segments: style.borderRightDashSegments,
		},
	];
}

/**
 * The dash choice of one edge line. `'solid'` removes it (no dash is solid);
 * `'custom'` without segments leaves the line's own dash as it is, so a
 * loaded `a:custDash` the parser could not read is not replaced.
 */
function writeEdgeDash(ln: XmlObject, edge: CellEdge): void {
	if (edge.dash === undefined || edge.dash === '') {
		return;
	}
	if (edge.dash === 'solid') {
		delete ln['a:prstDash'];
		delete ln['a:custDash'];
		return;
	}
	if (edge.dash === 'custom') {
		if (edge.segments && edge.segments.length > 0) {
			applyDrawingLineDash(ln, { strokeDash: 'custom', customDashSegments: edge.segments });
		}
		return;
	}
	applyDrawingLineDash(ln, { strokeDash: edge.dash as StrokeDashType });
}

/** Write each edge the style sets into `tcPr`, its children in `CT_LineProperties` order. */
export function writeCellEdgeBorders(
	tcPr: XmlObject,
	style: PptxTableCellStyle,
	emuPerPx: number,
): void {
	for (const edge of cellEdges(style)) {
		if (edge.width === undefined && edge.color === undefined && edge.dash === undefined) {
			continue;
		}
		const existing = tcPr[edge.xmlKey];
		const ln: XmlObject = existing && typeof existing === 'object' ? (existing as XmlObject) : {};
		if (edge.width !== undefined) {
			ln['@_w'] = String(Math.round(edge.width * emuPerPx));
		}
		if (edge.color) {
			delete ln['a:noFill'];
			delete ln['a:gradFill'];
			delete ln['a:pattFill'];
			ln['a:solidFill'] = {
				'a:srgbClr': { '@_val': edge.color.replace('#', '') },
			};
		}
		writeEdgeDash(ln, edge);
		tcPr[edge.xmlKey] = reorderObjectKeys(ln, LN_CHILD_ORDER);
	}
}
