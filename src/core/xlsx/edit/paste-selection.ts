import { MAX_COL, MAX_ROW, normalizeRange, type CellRange } from '../address';
import { cellsFromText, pasteAt } from './clipboard';
import type { EditContext } from './context';
import type { ClipboardPayload, PasteRequest } from './types';
import { resolvePasteOptions } from './paste-options';
import { pasteWidths } from './paste-widths';

/** Repeat a copied block over a compatible selection, using one undoable paste. */
export function pasteSelection(
	ctx: EditContext,
	sheet: number,
	selection: CellRange,
	payload: ClipboardPayload | string,
	request: PasteRequest,
): CellRange {
	const dest = normalizeRange(selection);
	const clip =
		typeof payload === 'string'
			? { tsv: payload, html: '', cells: cellsFromText(ctx.workbook, payload) }
			: payload;
	const options = resolvePasteOptions(request);
	if (options.mode === 'widths')
		return pasteWidths(ctx, sheet, dest, clip.cells, options.transpose, !!clip.cut);
	const height = options.transpose ? clip.cells.cols : clip.cells.rows;
	const width = options.transpose ? clip.cells.rows : clip.cells.cols;
	const rows = dest.end.row - dest.start.row + 1;
	const cols = dest.end.col - dest.start.col + 1;
	// A cut is one move. Full-sheet axes retain the existing bounded paste behavior.
	if (
		clip.cut ||
		!height ||
		!width ||
		(rows === 1 && cols === 1) ||
		(dest.start.row === 0 && dest.end.row === MAX_ROW) ||
		(dest.start.col === 0 && dest.end.col === MAX_COL)
	)
		return pasteAt(ctx, sheet, dest.start, clip, options);
	if (rows % height || cols % width)
		throw new RangeError('The copy and paste areas are not compatible sizes.');
	return pasteAt(ctx, sheet, dest.start, clip, options, dest);
}
