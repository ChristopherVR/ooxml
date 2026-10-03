// Undo history must record what an edit changes, not copy the sheet or the workbook per step.
import { describe, expect, it } from 'vitest';
import { parseRange } from '../address.js';
import { putCell } from '../cells.js';
import type { Workbook } from '../model.js';
import { createWorkbook } from '../workbook.js';
import { setCellInput } from './cell-values.js';
import { createEditSession } from './session.js';
import { insertRows } from './structure.js';
import { testContext } from './test-context.js';

function bigWorkbook(rows: number): Workbook {
	const wb = createWorkbook({ sheets: ['A', 'B'] });
	for (const sheet of wb.sheets)
		for (let r = 0; r < rows; r++)
			for (let c = 0; c < 10; c++) putCell(sheet, r, c, { value: r * c });
	return wb;
}

/** Cells held by a step's snapshots (the memory an undo step keeps alive). */
function recordedCells(entries: { before: object; after: object }[]): number {
	let n = 0;
	for (const e of entries)
		for (const snap of [e.before, e.after] as {
			kind: string;
			cells?: unknown[];
			data?: { formulas?: unknown[]; rows?: Map<number, Map<number, unknown>> };
		}[]) {
			n += snap.cells?.length ?? 0;
			n += snap.data?.formulas?.length ?? 0;
			for (const cells of snap.data?.rows?.values() ?? []) n += cells.size;
		}
	return n;
}

describe('undo history size and speed on large sheets', () => {
	it('20 table header edits on a 50k x 10 sheet take under a second and record a cell each', () => {
		const wb = bigWorkbook(50_000);
		const s = createEditSession(wb, { recalc: false, autoRowHeight: false });
		s.createTable(
			0,
			parseRange('A1:J100') ?? { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
			true,
		);
		const start = performance.now();
		for (let i = 0; i < 20; i++) s.setCellInput(0, 0, 0, `Head${i}`);
		expect(performance.now() - start).toBeLessThan(1000);
		const ctx = testContext(wb);
		setCellInput(ctx, 0, 0, 0, 'Next');
		expect(recordedCells(ctx.steps[0]?.entries ?? [])).toBe(2);
		ctx.undo();
		for (let i = 0; i < 20; i++) s.undo();
		expect(wb.sheets[0]?.tables[0]?.columns[0]?.name).toBe('0');
	});

	it('10 row inserts on a 500k-cell sheet take under two seconds and copy no moved cells', () => {
		const wb = bigWorkbook(50_000);
		const s = createEditSession(wb, { recalc: false, autoRowHeight: false });
		const start = performance.now();
		for (let i = 0; i < 10; i++) s.insertRows(1, 0, 1);
		expect(performance.now() - start).toBeLessThan(2000);
		const ctx = testContext(wb);
		insertRows(ctx, 1, 0, 1);
		// Only the inserted (empty) row's cells could be recorded; the 500k moved cells are not.
		expect(recordedCells(ctx.steps[0]?.entries ?? [])).toBeLessThanOrEqual(10);
		ctx.undo();
		for (let i = 0; i < 10; i++) s.undo();
		expect(wb.sheets[1]?.rows.get(49_999)?.get(9)?.value).toBe(49_999 * 9);
		expect(wb.sheets[1]?.rows.get(50_000)).toBeUndefined();
	});
});
