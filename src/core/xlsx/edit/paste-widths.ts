import { MAX_COL, MAX_ROW, type CellRange } from '../address.js';
import { DEFAULT_COL_WIDTH } from '../workbook.js';
import { editColumns } from './columns.js';
import { sheetAt, type EditContext } from './context.js';
import type { ClipboardCells } from './types.js';

/** Excel repeats original column metrics, even when the copied cells are transposed. */
export function pasteWidths(
	ctx: EditContext,
	s: number,
	selection: CellRange,
	cells: ClipboardCells,
	transpose: boolean,
	cut: boolean,
): CellRange {
	if (cut) throw new RangeError('Column widths are not available for cut cells.');
	const sheet = sheetAt(ctx.workbook, s);
	const widths = cells.columnWidths;
	if (
		!widths ||
		widths.length < cells.cols ||
		widths.some((w) => !Number.isFinite(w) || w < 0 || w > 255)
	)
		throw new RangeError('The clipboard does not contain valid column widths.');
	const sourceRows = cells.columnWidthRows ?? cells.rows;
	if (!Number.isInteger(sourceRows) || sourceRows < cells.rows || sourceRows > MAX_ROW + 1)
		throw new RangeError('The clipboard does not contain valid column widths.');
	const height = transpose ? widths.length : sourceRows;
	const width = transpose ? sourceRows : widths.length;
	if (!height || !width) return selection;
	const rows = selection.end.row - selection.start.row + 1;
	const cols = selection.end.col - selection.start.col + 1;
	// Unlike cell pastes, incompatible selections use one block without an error.
	const repeat = rows % height === 0 && cols % width === 0;
	const dest = repeat
		? selection
		: {
				start: selection.start,
				end: { row: selection.start.row + height - 1, col: selection.start.col + width - 1 },
			};
	if (dest.end.row > MAX_ROW || dest.end.col > MAX_COL)
		throw new RangeError('The paste area extends beyond the sheet.');
	const targets = Array.from(
		{ length: dest.end.col - dest.start.col + 1 },
		(_, c) => dest.start.col + c,
	);
	ctx.run(
		'Paste column widths',
		'view',
		[{ kind: 'parts', sheet: s, parts: ['columns'] }],
		() => {
			editColumns(sheet, targets, (info) => {
				const w = widths[(info.min - dest.start.col) % widths.length]!;
				info.width = w || (sheet.defaultColWidth ?? DEFAULT_COL_WIDTH);
				info.customWidth = true;
				delete info.bestFit;
				if (w === 0) info.hidden = true;
				else delete info.hidden;
			});
		},
		{ sheet: s, ranges: [dest], structural: true },
	);
	return dest;
}
