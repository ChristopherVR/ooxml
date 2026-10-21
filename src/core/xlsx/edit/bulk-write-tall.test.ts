// Blocks taller than a call's argument limit (about 125,000 in Node): writes, deferred
// recalculation and pasted text must not pass one argument per row or cell.
import { describe, expect, it } from 'vitest';
import { getCell } from '../cells';
import type { CellValue } from '../model';
import { createWorkbook } from '../workbook';
import { parseTsv } from './clipboard-text';
import { createEditSession } from './session';

describe('tall blocks', () => {
	const ROWS = 200_000;
	it('write a single column of any height', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook, { autoRowHeight: false });
		const values: CellValue[][] = Array.from({ length: ROWS }, (_, r) => [r]);
		session.setRangeValues(0, { row: 0, col: 0 }, values);
		expect(getCell(workbook.sheets[0]!, ROWS - 1, 0)?.value).toBe(ROWS - 1);
		session.undo();
		expect(workbook.sheets[0]!.rows.size).toBe(0);
	});

	it('defer recalculation in manual mode without passing each cell as an argument', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook, { autoRowHeight: false });
		session.setCellInput(0, 0, 1, '=SUM(A1:A200000)');
		session.setCalcMode('manual');
		session.setRangeValues(
			0,
			{ row: 0, col: 0 },
			Array.from({ length: ROWS }, () => [1]),
		);
		expect(getCell(workbook.sheets[0]!, 0, 1)?.value).toBe(0);
		session.calculateNow();
		expect(getCell(workbook.sheets[0]!, 0, 1)?.value).toBe(ROWS);
	});

	it('come from pasted text of any height', () => {
		const rows = parseTsv(Array.from({ length: ROWS }, (_, r) => `${r}\t${r}`).join('\n'));
		expect(rows).toHaveLength(ROWS);
		expect(rows.at(-1)).toEqual([`${ROWS - 1}`, `${ROWS - 1}`]);
	});
});
