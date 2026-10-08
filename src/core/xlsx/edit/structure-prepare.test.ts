// Structural edits that arrive while the formula graph is still being prepared in idle time (or
// before any preparation): the graph is finished for the workbook as it was, then moved in place,
// and the values equal a full recalculation of a copy from scratch.
import { describe, expect, it } from 'vitest';
import { getCell, putCell } from '../cells';
import { createCalcEngine } from '../formula/engine';
import type { CellValue, Workbook } from '../model';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import type { EditSession } from './types';

/** A workbook as a file loads it: legacy formulas whose stored values are up to date. */
function opened(): Workbook {
	const wb = createWorkbook({ sheets: ['Data', 'Report'] });
	const [data, report] = wb.sheets;
	if (!data || !report) throw new Error('sheets');
	const formula = (formula: string) => ({ value: null, formula, legacyFormula: true as const });
	for (let r = 0; r < 300; r++) {
		putCell(data, r, 0, { value: r * 3 });
		putCell(data, r, 1, formula(`A${r + 1}*2`));
		putCell(data, r, 2, formula(r === 0 ? 'B1' : `C${r}+B${r + 1}`));
		putCell(data, r, 3, formula(`ROW()+SUM($A$1:A${r + 1})`));
	}
	for (let r = 0; r < 20; r++) putCell(report, r, 0, formula(`Data!C${r * 10 + 1}+Data!A1:A300`));
	putCell(report, 0, 1, formula('SUM(Data!A:A)+INDEX(Data!B:B,7)'));
	createCalcEngine(wb).recalculateAll();
	return wb;
}

/** Values that differ from a full recalculation of a copy. */
function differences(workbook: Workbook): string[] {
	const copy = structuredClone(workbook);
	createCalcEngine(copy).recalculateAll();
	const out: string[] = [];
	workbook.sheets.forEach((sheet, s) => {
		for (const [row, cells] of copy.sheets[s]?.rows ?? [])
			for (const [col, want] of cells) {
				const got: CellValue = getCell(sheet, row, col)?.value ?? null;
				if (JSON.stringify(got) !== JSON.stringify(want.value))
					out.push(
						`${sheet.name}!${row},${col}: ${JSON.stringify(got)} vs ${JSON.stringify(want.value)}`,
					);
			}
	});
	return out;
}

/** Counts the graph rebuilds (`invalidate`) a session's engine goes through. */
function countRebuilds(session: EditSession): () => number {
	const calc = session.calc;
	const invalidate = calc.invalidate.bind(calc);
	let count = 0;
	calc.invalidate = () => {
		count++;
		invalidate();
	};
	return () => count;
}

const spent = () => 0;

describe('structural edits and graph preparation', () => {
	for (const [label, slices] of [
		['before any preparation', 0],
		['while preparation is half done', 10],
		['after preparation finished', Number.POSITIVE_INFINITY],
	] as const)
		it(`follow an edit made ${label}`, () => {
			const wb = opened();
			const session = createEditSession(wb, { autoRowHeight: false });
			const rebuilds = countRebuilds(session);
			for (let i = 0; i < slices; i++)
				if (session.prepareCalculation({ timeRemaining: spent })) break;
			session.insertRows(0, 5, 2);
			expect(differences(wb)).toEqual([]);
			session.deleteColumns(0, 0, 1);
			expect(differences(wb)).toEqual([]);
			session.renameSheet(0, 'Inputs');
			expect(differences(wb)).toEqual([]);
			session.setCellInput(0, 10, 0, '7');
			expect(differences(wb)).toEqual([]);
			expect(rebuilds()).toBe(0);
		});

	it('follows an edit made in manual mode while preparation is half done', () => {
		const wb = opened();
		wb.calcMode = 'manual';
		const session = createEditSession(wb, { autoRowHeight: false });
		session.prepareCalculation({ timeRemaining: spent });
		session.insertRows(0, 5, 2);
		session.setCellInput(0, 3, 0, '100');
		session.calculateNow();
		expect(differences(wb)).toEqual([]);
	});
});
