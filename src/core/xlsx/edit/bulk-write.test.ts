// Bulk writes through the edit session: their undo records stay cheap per cell, and undo, redo and
// recalculation give exactly the results of a full copy of the sheet.
import { afterEach, describe, expect, it, vi } from 'vitest';
import { getCell, putCell } from '../cells';
import type { CellValue, Workbook } from '../model';
import { internStyle } from '../styles';
import { createWorkbook, defaultCellStyle } from '../workbook';
import { createEditSession } from './session';

afterEach(() => {
	vi.restoreAllMocks();
});

const COLS = 10;

/** A block of plain values; `''` and `null` clear cells. */
function block(rows: number, salt: number): CellValue[][] {
	return Array.from({ length: rows }, (_, r) =>
		Array.from({ length: COLS }, (_, c) => {
			const n = r * COLS + c + salt;
			if (n % 17 === 0) return null;
			if (n % 13 === 0) return '';
			if (n % 5 === 0) return n % 2 === 0;
			return c % 3 === 0 ? `item ${n}` : n + 0.25;
		}),
	);
}

/** A sheet holding every kind of cell content the writes replace or keep. */
function mixedWorkbook(rows: number): Workbook {
	const wb = createWorkbook();
	const sheet = wb.sheets[0];
	if (!sheet) throw new Error('no sheet');
	const base = defaultCellStyle();
	const bold = internStyle(wb, { ...base, font: { ...base.font, bold: true } });
	for (let r = 0; r < rows; r += 2)
		for (let c = 0; c < COLS; c += 3) putCell(sheet, r, c, { value: r * c, styleId: bold });
	putCell(sheet, 1, 1, {
		value: 'rich text',
		richText: [{ text: 'rich ', font: { bold: true } }, { text: 'text' }],
	});
	putCell(sheet, 3, 2, { value: { error: '#N/A' } });
	putCell(sheet, 5, 4, {
		value: 6,
		formula: 'A1:B2*2',
		arrayRange: { start: { row: 5, col: 4 }, end: { row: 6, col: 5 } },
	});
	putCell(sheet, 7, 6, { value: null, styleId: bold });
	return wb;
}

/** How many times `structuredClone` runs while a session writes, undoes and redoes `rows` rows. */
function clonesFor(rows: number): number {
	const session = createEditSession(mixedWorkbook(rows), { recalc: false, autoRowHeight: false });
	const spy = vi.spyOn(globalThis, 'structuredClone');
	session.setRangeValues(0, { row: 0, col: 0 }, block(rows, 1));
	session.undo();
	session.redo();
	const calls = spy.mock.calls.length;
	spy.mockRestore();
	return calls;
}

describe('bulk write undo records', () => {
	it('do not deep-clone each plain cell', () => {
		expect(clonesFor(2000)).toBe(clonesFor(200));
	});

	it('are copies: later changes to a cell do not reach them', () => {
		const wb = mixedWorkbook(20);
		const sheet = wb.sheets[0]!;
		const session = createEditSession(wb, { recalc: false, autoRowHeight: false });
		const before = structuredClone(sheet);
		session.setRangeValues(0, { row: 0, col: 0 }, block(20, 3));
		const after = structuredClone(sheet);
		// Mutate cells in place (as an outside caller could) after the step was recorded.
		putCell(sheet, 1, 1, {
			value: 'rich text',
			richText: [{ text: 'rich ', font: { bold: true } }, { text: 'text' }],
		});
		getCell(sheet, 1, 1)!.richText!.push({ text: ' changed' });
		session.undo();
		expect(sheet).toStrictEqual(before);
		getCell(sheet, 1, 1)!.richText![0]!.text = 'changed';
		(getCell(sheet, 5, 4)!.arrayRange!.end as { row: number }).row = 99;
		session.redo();
		expect(sheet).toStrictEqual(after);
	});
});

describe('bulk write undo and redo', () => {
	it('restore every state exactly, over mixed and overlapping writes', () => {
		const wb = mixedWorkbook(300);
		const sheet = wb.sheets[0]!;
		const session = createEditSession(wb, { recalc: false, autoRowHeight: false });
		const steps: (() => void)[] = [
			() => session.setRangeValues(0, { row: 0, col: 0 }, block(300, 1)),
			() => session.setRangeValues(0, { row: 150, col: 4 }, block(300, 2)),
			() => session.clearRange(0, { start: { row: 10, col: 0 }, end: { row: 60, col: 5 } }, 'all'),
			// Overlapping ranges in one step: each cell is recorded once.
			() =>
				session.applyStyle(
					0,
					[
						{ start: { row: 0, col: 0 }, end: { row: 40, col: 6 } },
						{ start: { row: 20, col: 3 }, end: { row: 80, col: 12 } },
					],
					{ font: { italic: true } },
				),
			() => {
				const payload = session.copy(0, { start: { row: 0, col: 0 }, end: { row: 99, col: 9 } });
				session.paste(0, { row: 200, col: 2 }, payload);
			},
			() => session.setCellInput(0, 450, 3, 'typed'),
		];
		const states = [structuredClone(sheet)];
		for (const step of steps) {
			step();
			states.push(structuredClone(sheet));
		}
		for (let i = steps.length - 1; i >= 0; i--) {
			session.undo();
			expect(sheet).toStrictEqual(states[i]);
		}
		for (let i = 1; i <= steps.length; i++) {
			session.redo();
			expect(sheet).toStrictEqual(states[i]);
		}
	});
});

describe('bulk write recalculation', () => {
	it('recalculates formulas that depend on the written cells, also on undo and redo', () => {
		const rows = 500;
		const wb = createWorkbook();
		const sheet = wb.sheets[0]!;
		for (let r = 0; r < rows; r++) {
			putCell(sheet, r, 0, { value: r });
			putCell(sheet, r, 2, { value: null, formula: `A${r + 1}*2` });
		}
		putCell(sheet, 0, 3, { value: null, formula: `SUM(A1:A${rows})` });
		const session = createEditSession(wb, { autoRowHeight: false });
		session.calculateNow();
		const column = (values: (r: number) => number) =>
			session.setRangeValues(
				0,
				{ row: 0, col: 0 },
				Array.from({ length: rows }, (_, r) => [values(r)]),
			);
		const expectSums = (value: (r: number) => number) => {
			for (const r of [0, 1, rows / 2, rows - 1])
				expect(getCell(sheet, r, 2)?.value).toBe(value(r) * 2);
			let total = 0;
			for (let r = 0; r < rows; r++) total += value(r);
			expect(getCell(sheet, 0, 3)?.value).toBe(total);
		};
		column((r) => r * 3 + 1);
		expectSums((r) => r * 3 + 1);
		session.undo();
		expectSums((r) => r);
		session.redo();
		expectSums((r) => r * 3 + 1);
	});
});
