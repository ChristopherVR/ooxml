/**
 * table-cell-paragraphs.ts - read each table cell paragraph's own layout.
 *
 * Reads each `a:p`'s own `a:pPr` layout (the fields shape text reads) into
 * `PptxTableCell.paragraphs`. `PptxTableCell.style` still holds only the first
 * paragraph's alignment.
 *
 * @module table-cell-paragraphs
 */
import type { PptxTableCellParagraph, XmlObject } from '../../types';
import {
	parseAlignmentAttr,
	parseLineSpacingExactPt,
	parseLineSpacingMultiplier,
	parseParagraphMargins,
	parseParagraphRtl,
	parseParagraphSpacingPx,
} from '../../utils/paragraph-properties-parser';
import type { TableCellRunsContext } from './table-cell-runs';

const PX_PER_POINT = 96 / 72;

/** `a:rPr@sz` (hundredths of a point) in points, when it is set. */
function sizePt(runProperties: unknown): number | undefined {
	const size = Number.parseInt(
		String((runProperties as XmlObject | undefined)?.['@_sz'] ?? ''),
		10,
	);
	return Number.isFinite(size) && size > 0 ? size / 100 : undefined;
}

function readParagraph(
	paragraph: XmlObject,
	context: Pick<TableCellRunsContext, 'ensureArray'>,
	defaultFontSizePt: number | undefined,
): PptxTableCellParagraph {
	const result: PptxTableCellParagraph = {};
	const firstRun = context.ensureArray(paragraph['a:r'])[0] as XmlObject | undefined;
	const isEmpty =
		!firstRun &&
		context.ensureArray(paragraph['a:fld']).length === 0 &&
		context.ensureArray(paragraph['a:br']).length === 0;
	const endSizePt = sizePt(paragraph['a:endParaRPr']);
	if (isEmpty && endSizePt !== undefined) {
		result.endParaFontSize = endSizePt;
	}

	const pPr = paragraph['a:pPr'] as XmlObject | undefined;
	if (!pPr || typeof pPr !== 'object') {
		return result;
	}
	const align = parseAlignmentAttr(pPr['@_algn'] as string | undefined);
	if (align) {
		result.align = align;
	}
	const rtl = parseParagraphRtl(pPr);
	if (rtl !== undefined) {
		result.rtl = rtl;
	}
	const { paragraphMarginLeft, paragraphIndent } = parseParagraphMargins(pPr);
	if (paragraphMarginLeft !== undefined) {
		result.paragraphMarginLeft = paragraphMarginLeft;
	}
	if (paragraphIndent !== undefined) {
		result.paragraphIndent = paragraphIndent;
	}
	const lnSpc = pPr['a:lnSpc'] as XmlObject | undefined;
	const lineSpacing = parseLineSpacingMultiplier(lnSpc);
	const exactPt = parseLineSpacingExactPt(lnSpc);
	if (lineSpacing !== undefined) {
		result.lineSpacing = lineSpacing;
	} else if (exactPt !== undefined) {
		result.lineSpacingExactPt = exactPt;
	}
	// Percentage spacing resolves against the paragraph's own size: its first
	// run's, its end properties' when it has no run, or the table's default.
	const basisPt = (firstRun ? sizePt(firstRun['a:rPr']) : endSizePt) ?? defaultFontSizePt;
	const basisPx = basisPt === undefined ? undefined : basisPt * PX_PER_POINT;
	const before = parseParagraphSpacingPx(pPr['a:spcBef'] as XmlObject | undefined, basisPx);
	if (before !== undefined) {
		result.paragraphSpacingBefore = before;
	}
	const after = parseParagraphSpacingPx(pPr['a:spcAft'] as XmlObject | undefined, basisPx);
	if (after !== undefined) {
		result.paragraphSpacingAfter = after;
	}
	return result;
}

/**
 * Read each paragraph's own layout in a table cell.
 *
 * @param tableCell - The `a:tc` node.
 * @param context - The parser context (only `ensureArray` is used).
 * @param defaultFontSizePt - The table's default text size, for percentage
 *   spacing in paragraphs that set no size of their own.
 * @returns One entry per `a:p`, or `undefined` when no paragraph sets any of
 *   it, so an unformatted cell keeps its existing shape.
 */
export function extractTableCellParagraphs(
	tableCell: XmlObject | undefined,
	context: Pick<TableCellRunsContext, 'ensureArray'>,
	defaultFontSizePt?: number,
): PptxTableCellParagraph[] | undefined {
	const paragraphs = context.ensureArray(
		(tableCell?.['a:txBody'] as XmlObject | undefined)?.['a:p'],
	) as XmlObject[];
	const read = paragraphs.map((paragraph) =>
		readParagraph(paragraph ?? {}, context, defaultFontSizePt),
	);
	return read.some((entry) => Object.keys(entry).length > 0) ? read : undefined;
}
