import type { Workbook, Worksheet } from '../model.js';
import { cellView } from './cell-view.js';
import type { FontView } from './types.js';
import {
	COLUMN_PADDING_PX,
	DEFAULT_MAX_DIGIT_WIDTH,
	defaultColumnPixels,
	pixelsToColumnWidth,
} from './units.js';

/** Measures `text` in `font`, in CSS pixels at 100% zoom (injected: canvas, DOM or a font table). */
export type MeasureText = (text: string, font: FontView) => number;

/**
 * The file width (as `ColumnInfo.width`) that fits every cell in column `col`, like Excel's
 * AutoFit: the widest display text plus the 5 px cell padding and the 1 px gridline. Cells inside merges spanning several
 * columns are ignored, as Excel does; multi-line text uses its longest line. A column without
 * text gets the sheet's default width.
 */
export function autoFitColumnWidth(
	sheet: Worksheet | number,
	workbook: Workbook,
	col: number,
	measure: MeasureText,
	mdw = DEFAULT_MAX_DIGIT_WIDTH,
): number {
	const sheetIndex = typeof sheet === 'number' ? sheet : workbook.sheets.indexOf(sheet);
	const ws = workbook.sheets[sheetIndex];
	if (!ws) return pixelsToColumnWidth(defaultColumnPixels(undefined, mdw), mdw);
	let widest = 0;
	for (const [row, cells] of ws.rows) {
		if (!cells.has(col)) continue;
		const merged = ws.merges.some(
			(m) =>
				row >= m.start.row &&
				row <= m.end.row &&
				col >= m.start.col &&
				col <= m.end.col &&
				m.start.col !== m.end.col,
		);
		if (merged) continue;
		const view = cellView(workbook, sheetIndex, row, col);
		if (!view.text) continue;
		let width: number;
		if (view.rich?.length) {
			width = view.rich.reduce((sum, run) => sum + measure(run.text, run.font), 0);
		} else {
			width = Math.max(...view.text.split(/\r?\n/).map((line) => measure(line, view.font)));
		}
		widest = Math.max(widest, width + view.indentPx);
	}
	if (widest <= 0)
		return ws.defaultColWidth ?? pixelsToColumnWidth(defaultColumnPixels(undefined, mdw), mdw);
	return pixelsToColumnWidth(Math.ceil(widest) + COLUMN_PADDING_PX + 1, mdw);
}
