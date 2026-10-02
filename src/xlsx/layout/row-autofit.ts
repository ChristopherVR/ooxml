// Automatic row height: the tallest line stack of a row's cells, like Excel's AutoFit.
import type { Workbook } from '../model.js';
import type { MeasureText } from './autofit.js';
import { cellView } from './cell-view.js';
import { createGridMetrics } from './metrics.js';
import { COLUMN_PADDING_PX } from './units.js';

/** Excel's row heights (points) for Calibri at common font sizes (points). */
const LINE_HEIGHTS: readonly [number, number][] = [
	[6, 8.25],
	[8, 11.25],
	[9, 12],
	[10, 12.75],
	[11, 15],
	[12, 15.75],
	[14, 18.75],
	[16, 21],
	[18, 23.25],
	[20, 26.25],
	[22, 28.5],
	[24, 31.5],
	[26, 33.75],
	[28, 36.75],
	[36, 46.5],
	[48, 63],
	[72, 93.75],
];

/** Points one line of text in a font of `sizePt` needs (interpolated, whole pixels). */
export function lineHeightPoints(sizePt: number): number {
	const first = LINE_HEIGHTS[0] as [number, number];
	const last = LINE_HEIGHTS[LINE_HEIGHTS.length - 1] as [number, number];
	let height: number;
	if (sizePt <= first[0]) height = (first[1] / first[0]) * sizePt;
	else if (sizePt >= last[0]) height = (last[1] / last[0]) * sizePt;
	else {
		let i = 1;
		while ((LINE_HEIGHTS[i] as [number, number])[0] < sizePt) i++;
		const [s0, h0] = LINE_HEIGHTS[i - 1] as [number, number];
		const [s1, h1] = LINE_HEIGHTS[i] as [number, number];
		height = h0 + ((h1 - h0) * (sizePt - s0)) / (s1 - s0);
	}
	// Row heights are whole pixels (0.75 pt at 96 dpi).
	return Math.max(0.75, Math.ceil(height / 0.75 - 1e-9) * 0.75);
}

/** How many lines `text` wraps into within `width` px (word wrap, explicit line breaks kept). */
export function wrappedLineCount(
	text: string,
	width: number,
	measure: (text: string) => number,
): number {
	let lines = 0;
	for (const paragraph of text.split(/\r?\n/)) {
		lines++;
		if (width <= 0) continue;
		let current = '';
		for (const word of paragraph.split(/(\s+)/)) {
			if (word === '') continue;
			const next = current + word;
			if (current.trim() === '' || measure(next) <= width) current = next;
			else {
				lines++;
				current = word.trimStart();
			}
			// A single word wider than the cell breaks by character.
			while (current.length > 1 && measure(current) > width) {
				let fit = 1;
				while (fit < current.length && measure(current.slice(0, fit + 1)) <= width) fit++;
				current = current.slice(fit);
				lines++;
			}
		}
	}
	return lines;
}

/**
 * The height in points that fits every cell of `row`, as Excel's row AutoFit computes it:
 * the largest font's line height, times the wrapped line count for cells with wrap text
 * (measured against the column width less the cell padding). Cells inside merges are ignored,
 * as Excel does. A row without text gets the sheet's default height. Line heights follow
 * Calibri metrics, so other fonts are approximate.
 */
export function autoFitRowHeight(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	measure: MeasureText,
): number {
	const sheet = workbook.sheets[sheetIndex];
	if (!sheet) return 15;
	const cells = sheet.rows.get(row);
	if (!cells?.size) return sheet.defaultRowHeight;
	const metrics = createGridMetrics(sheet, { zoom: 100 });
	let tallest = 0;
	for (const col of cells.keys()) {
		const merged = sheet.merges.some(
			(m) =>
				row >= m.start.row &&
				row <= m.end.row &&
				col >= m.start.col &&
				col <= m.end.col &&
				(m.start.row !== m.end.row || m.start.col !== m.end.col),
		);
		if (merged) continue;
		const view = cellView(workbook, sheetIndex, row, col);
		if (!view.text) continue;
		const fonts = view.rich?.length ? view.rich.map((run) => run.font) : [view.font];
		const sizePt = Math.max(...fonts.map((f) => (f.sizePx * 3) / 4));
		let lines = 1;
		if (view.wrap) {
			const width = metrics.colWidth(col) - COLUMN_PADDING_PX - 1 - view.indentPx;
			lines = wrappedLineCount(view.text, width, (text) => measure(text, view.font));
		}
		tallest = Math.max(tallest, lines * lineHeightPoints(sizePt));
	}
	return tallest > 0 ? tallest : sheet.defaultRowHeight;
}

/** A font-size based text measure (px) for when no real text measurement is available. */
export const approximateMeasure: MeasureText = (text, font) =>
	text.length * font.sizePx * (font.bold ? 0.58 : 0.53);
