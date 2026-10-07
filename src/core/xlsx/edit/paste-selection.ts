import { MAX_COL, MAX_ROW, normalizeRange, type CellRange } from '../address.js';
import { cellsFromText, pasteAt } from './clipboard.js';
import type { EditContext } from './context.js';
import type { ClipboardPayload, PasteRequest } from './types.js';
import { resolvePasteOptions } from './paste-options.js';
import { pasteWidths } from './paste-widths.js';

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
	if (rows * cols > 250_000) throw new RangeError('The selected paste area is too large.');
	// The outer step captures every destination cell and merge once, rather than per tile.
	const tileContext: EditContext = { ...ctx, run: (_label, _kind, _scopes, fn) => fn() };
	return ctx.run(
		'Paste',
		'cells',
		[
			{ kind: 'cells', sheet, ranges: [dest] },
			{ kind: 'parts', sheet, parts: ['merges', 'comments', 'dataValidations'] },
		],
		() => {
			for (let row = dest.start.row; row <= dest.end.row; row += height)
				for (let col = dest.start.col; col <= dest.end.col; col += width)
					pasteAt(tileContext, sheet, { row, col }, clip, options);
			return dest;
		},
		{ sheet, ranges: [dest] },
	);
}
