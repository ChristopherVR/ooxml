import { describe, expect, it } from 'vitest';
import { formatRange } from '../address';
import { getCell, putCell } from '../cells';
import type { Workbook } from '../model';
import { internStyle, patchStyle } from '../styles';
import { createWorkbook, defaultCellStyle } from '../workbook';
import { compareWorkbooks } from './compare';

const ROWS = 1000;
const COLS = 250;

/** A dense numeric sheet; `reverse` stores the rows (and cells) in the opposite order. */
function grid(
	edit?: (book: Workbook, row: number, col: number) => number | undefined,
	reverse = false,
) {
	const book = createWorkbook();
	const sheet = book.sheets[0]!;
	for (let i = 0; i < ROWS; i++) {
		const row = reverse ? ROWS - 1 - i : i;
		for (let j = 0; j < COLS; j++) {
			const col = reverse ? COLS - 1 - j : j;
			const value = row * COLS + col;
			const styleId = edit?.(book, row, col);
			putCell(sheet, row, col, styleId === undefined ? { value } : { value, styleId });
		}
	}
	return book;
}

describe('compareWorkbooks on a large sheet', () => {
	it('compares 250,000 cells per side and reports exactly the changed ones', () => {
		const before = grid();
		expect(compareWorkbooks(before, grid(undefined, true)).count).toBe(0);

		let bold: number | undefined;
		const after = grid((book, row, col) => {
			bold ??= internStyle(book, patchStyle(defaultCellStyle(), { font: { bold: true } }));
			return row >= 500 && row < 600 && col >= 100 && col < 110 ? bold : undefined;
		}, true);
		const sheet = after.sheets[0]!;
		for (let row = 0; row < ROWS; row += 100)
			for (let col = 0; col < COLS; col += 50)
				putCell(sheet, row, col, { ...getCell(sheet, row, col), value: -1 });

		const result = compareWorkbooks(before, after);
		const { cells, ranges } = result.sheets[0]!;
		expect(result.count).toBe(1050);
		expect(cells.filter((d) => d.kind === 'enteredValue')).toHaveLength(50);
		expect(cells.filter((d) => d.kind === 'format')).toHaveLength(1000);
		expect(ranges).toHaveLength(51);
		expect(ranges.filter((r) => r.count > 1).map((r) => [formatRange(r.range), r.count])).toEqual([
			['CW501:DF600', 1000],
		]);
		// Long structural lists neither overflow the call stack nor escape the cap.
		const merged = createWorkbook();
		merged.sheets[0]!.merges = Array.from({ length: 200_000 }, (_, row) => ({
			start: { row, col: 0 },
			end: { row, col: 1 },
		}));
		expect(compareWorkbooks(createWorkbook(), merged).count).toBe(200_000);
		expect(compareWorkbooks(createWorkbook(), merged, { maxDifferences: 10 })).toMatchObject({
			count: 10,
			truncated: true,
		});
		// Row-major order whatever order the cells were stored in.
		const keys = cells.map((d) => d.address.row * COLS + d.address.col);
		expect(keys).toEqual([...keys].sort((a, b) => a - b));
	});
});
