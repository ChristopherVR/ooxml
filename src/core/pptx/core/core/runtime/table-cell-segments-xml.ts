/**
 * A table cell's `a:txBody` written from rich segments
 * (`PptxTableCell.textSegments`), for cells made in code.
 *
 * @module runtime/table-cell-segments-xml
 */

import type { PptxTableCell, TextSegment, XmlObject } from '../../types';

/** The text a cell's segments spell, as the flat `PptxTableCell.text` writes it. */
export function tableCellSegmentsText(segments: readonly TextSegment[]): string {
	return segments
		.map((segment) => (segment.isParagraphBreak ? '\n' : (segment.text ?? '')))
		.join('');
}

/** The segments to write for a cell, or `undefined` when its `text` was changed since. */
export function tableCellSegmentsForSave(cell: PptxTableCell): TextSegment[] | undefined {
	const segments = cell.textSegments;
	if (!segments || segments.length === 0) {
		return undefined;
	}
	return tableCellSegmentsText(segments) === cell.text ? segments : undefined;
}

/**
 * A cell text body around `paragraphs`, keeping the existing `a:bodyPr` and
 * `a:lstStyle`, in CT_TextBody order (`bodyPr`, `lstStyle?`, `p+`).
 */
export function buildCellTextBody(
	existingTxBody: XmlObject | undefined,
	paragraphs: XmlObject[],
): XmlObject {
	const existing =
		existingTxBody && typeof existingTxBody === 'object' ? existingTxBody : undefined;
	const txBody: XmlObject = {};
	txBody['a:bodyPr'] = existing?.['a:bodyPr'] ?? {};
	txBody['a:lstStyle'] = existing?.['a:lstStyle'] ?? {};
	txBody['a:p'] = paragraphs.length === 1 ? paragraphs[0] : paragraphs;
	return txBody;
}
