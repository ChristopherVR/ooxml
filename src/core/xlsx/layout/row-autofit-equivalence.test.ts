// The batched row fit must give exactly the heights of the per-row computation it replaced.
import { describe, expect, it } from 'vitest';
import { putCell } from '../cells';
import type { Font, Workbook } from '../model';
import { internStyle } from '../styles';
import { createWorkbook, defaultCellStyle } from '../workbook';
import type { MeasureText } from './autofit';
import { cellView } from './cell-view';
import { createGridMetrics } from './metrics';
import {
	approximateMeasure,
	autoFitRowHeight,
	fitRowHeights,
	lineHeightPoints,
	wrappedLineCount,
} from './row-autofit';
import { COLUMN_PADDING_PX } from './units';

/** The former `autoFitRowHeight`: full grid metrics and a scan of every merge, per row. */
function referenceRowHeight(
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

const WORDS = ['alpha', 'be', 'gamma', 'delta-epsilon', 'z', 'supercalifragilistic', 'x y', 'q'];
const ROWS = 120;
const COLS = 10;

/** A sheet mixing wrapped text, rich runs, font sizes, indents, merges and custom columns. */
function syntheticWorkbook(): Workbook {
	const wb = createWorkbook();
	const sheet = wb.sheets[0];
	if (!sheet) throw new Error('no sheet');
	const base = defaultCellStyle();
	const font = (size: number, bold = false): Font => ({ ...base.font, size, bold });
	const styles = [
		0,
		internStyle(wb, { ...base, font: font(8) }),
		internStyle(wb, { ...base, font: font(18, true) }),
		internStyle(wb, { ...base, alignment: { wrapText: true } }),
		internStyle(wb, {
			...base,
			font: font(14),
			alignment: { wrapText: true, horizontal: 'left', indent: 2 },
		}),
		internStyle(wb, { ...base, alignment: { horizontal: 'justify' } }),
		internStyle(wb, {
			...base,
			font: font(26),
			alignment: { wrapText: true, horizontal: 'right', indent: 3 },
		}),
		internStyle(wb, { ...base, alignment: { horizontal: 'distributed', indent: 1 } }),
	];
	const wrapped = styles[3] ?? 0;
	sheet.view.zoom = 150;
	sheet.defaultColWidth = 11.5;
	sheet.columns.push(
		{ min: 1, max: 1, width: 3.5, customWidth: true },
		{ min: 3, max: 4, width: 22 },
		{ min: 5, max: 5, hidden: true },
		{ min: 7, max: 9, width: 0.5 },
	);
	for (let row = 0; row < ROWS; row++) {
		if (row % 11 === 10) continue;
		for (let col = 0; col < COLS; col++) {
			if ((row + col) % 4 === 3) continue;
			const n = (row * 7 + col * 3) % WORDS.length;
			const words = Array.from(
				{ length: 1 + ((row + col) % 9) },
				(_, i) => WORDS[(n + i) % WORDS.length],
			);
			const text = words.join((row + col) % 5 === 0 ? '\n' : ' ');
			const value =
				col === 6 ? row * 1234.5678 : col === 8 ? row % 2 === 0 : (row + col) % 13 ? text : '';
			putCell(sheet, row, col, { value, styleId: styles[(row * 3 + col) % styles.length] ?? 0 });
		}
		if (row % 6 === 1)
			putCell(sheet, row, 2, {
				value: 'small BIG tail end',
				styleId: wrapped,
				richText: [
					{ text: 'small ', font: font(9) },
					{ text: 'BIG', font: font(20, true) },
					{ text: ' tail end' },
				],
			});
	}
	// Cells that decide their row's height unless a merge spanning several cells hides them.
	for (const [row, col] of [
		[30, 1],
		[20, 0],
		[13, 2],
		[41, 5],
		[70, 0],
	] as const)
		putCell(sheet, row, col, { value: 'tall text', styleId: styles[6] ?? 0 });
	// Rows the loop leaves empty, each holding one large cell: in a single-cell merge it counts,
	// in a wider merge it is ignored.
	const huge = internStyle(wb, { ...base, font: font(48) });
	putCell(sheet, 32, 1, { value: 'counted', styleId: huge });
	putCell(sheet, 43, 1, { value: 'ignored', styleId: huge });
	sheet.merges.push(
		{ start: { row: 12, col: 2 }, end: { row: 15, col: 2 } },
		{ start: { row: 20, col: 0 }, end: { row: 20, col: 3 } },
		{ start: { row: 32, col: 1 }, end: { row: 32, col: 1 } },
		{ start: { row: 43, col: 0 }, end: { row: 43, col: 2 } },
		{ start: { row: 40, col: 4 }, end: { row: 41, col: 6 } },
		{ start: { row: 50, col: 0 }, end: { row: 90, col: 0 } },
		{ start: { row: 500, col: 0 }, end: { row: 501, col: 1 } },
		// Overlapping merges only come from a malformed file: the spanning one still hides the cell.
		{ start: { row: 60, col: 5 }, end: { row: 62, col: 7 } },
		{ start: { row: 61, col: 6 }, end: { row: 61, col: 6 } },
	);
	putCell(sheet, 61, 6, { value: 'hidden by the spanning merge', styleId: huge });
	sheet.rowInfo.set(5, { height: 40, customHeight: true });
	sheet.rowInfo.set(6, { hidden: true });
	sheet.rowInfo.set(7, { height: 9 });
	return wb;
}

/** A measure that weighs characters unevenly and depends on the font. */
const unevenMeasure: MeasureText = (text, font) =>
	[...text].reduce(
		(width, ch) => width + (ch === ' ' ? 0.3 : /[A-Z]/.test(ch) ? 0.72 : 0.55) * font.sizePx,
		0,
	) * (font.bold ? 1.08 : 1);

const MEASURES: [string, MeasureText][] = [
	['approximate', approximateMeasure],
	['uneven', unevenMeasure],
];

describe('fitRowHeights', () => {
	const rows = Array.from({ length: ROWS + 10 }, (_, row) => row);

	it.each(MEASURES)('matches the per-row computation (%s measure)', (_, measure) => {
		const wb = syntheticWorkbook();
		const expected = rows.map((row) => referenceRowHeight(wb, 0, row, measure));
		// The sheet really exercises the fit: many distinct heights, not just the default.
		expect(new Set(expected).size).toBeGreaterThan(5);
		const batch = fitRowHeights(wb, 0, rows, measure);
		expect(rows.map((row) => batch.get(row))).toEqual(expected);
		expect(rows.map((row) => autoFitRowHeight(wb, 0, row, measure))).toEqual(expected);
	});

	it('fits any subset of rows like the whole sheet', () => {
		const wb = syntheticWorkbook();
		const subset = [91, 13, 40, 13, 50, 2, 30, 20];
		const batch = fitRowHeights(wb, 0, subset, approximateMeasure);
		expect([...batch.keys()].sort((a, b) => a - b)).toEqual([2, 13, 20, 30, 40, 50, 91]);
		for (const row of subset)
			expect(batch.get(row)).toBe(referenceRowHeight(wb, 0, row, approximateMeasure));
	});

	it('gives an empty map for a missing sheet', () => {
		const wb = syntheticWorkbook();
		expect(fitRowHeights(wb, 3, [0, 1], approximateMeasure).size).toBe(0);
		expect(autoFitRowHeight(wb, 3, 0, approximateMeasure)).toBe(15);
	});
});
