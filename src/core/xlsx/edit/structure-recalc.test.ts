// Row, column and sheet edits recalculate in place (the graph moves instead of being rebuilt, and
// only formulas whose value can change are evaluated). The oracle: after every edit, the values
// equal a full recalculation of a copy of the workbook from scratch.
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { getCell, putCell } from '../cells';
import { createCalcEngine } from '../formula/engine';
import type { Cell, CellValue, Workbook } from '../model';
import { loadXlsx } from '../read/load';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import type { EditSession } from './types';

const VOLATILE = /\b(?:RAND|RANDBETWEEN|RANDARRAY|NOW|TODAY)\s*\(/i;

const same = (a: CellValue, b: CellValue, tolerance: number): boolean => {
	if (typeof a === 'number' && typeof b === 'number')
		return Math.abs(a - b) <= tolerance * Math.max(1, Math.abs(a), Math.abs(b));
	return JSON.stringify(a) === JSON.stringify(b);
};

/** Values that differ from a full recalculation of a copy, as `Sheet!row,col: got vs want`. */
function differences(workbook: Workbook, tolerance = 1e-9): string[] {
	const copy = structuredClone(workbook);
	createCalcEngine(copy).recalculateAll();
	const out: string[] = [];
	workbook.sheets.forEach((sheet, s) => {
		const other = copy.sheets[s];
		const keys = new Set<string>();
		for (const [row, cells] of sheet.rows)
			for (const col of cells.keys()) keys.add(`${row},${col}`);
		for (const [row, cells] of other?.rows ?? [])
			for (const col of cells.keys()) keys.add(`${row},${col}`);
		for (const key of keys) {
			const [row, col] = key.split(',').map(Number) as [number, number];
			const got = getCell(sheet, row, col);
			const want = other && getCell(other, row, col);
			if (VOLATILE.test(got?.formula ?? '') || VOLATILE.test(want?.formula ?? '')) continue;
			if (!same(got?.value ?? null, want?.value ?? null, tolerance))
				out.push(
					`${sheet.name}!${key}: ${JSON.stringify(got?.value)} vs ${JSON.stringify(want?.value)}`,
				);
		}
	});
	return out;
}

type Edit = [string, (session: EditSession, sheet: number) => void];
const undo: Edit = ['undo', (s) => s.undo()];
const redo: Edit = ['redo', (s) => s.redo()];
const EDITS: Edit[] = [
	['insert a row at the top', (s, i) => s.insertRows(i, 0, 1)],
	undo,
	redo,
	['insert rows inside', (s, i) => s.insertRows(i, 3, 2)],
	['delete a row', (s, i) => s.deleteRows(i, 2, 1)],
	undo,
	redo,
	['delete rows inside', (s, i) => s.deleteRows(i, 4, 3)],
	['insert a column', (s, i) => s.insertColumns(i, 1, 1)],
	['delete a column', (s, i) => s.deleteColumns(i, 0, 1)],
	['insert columns far right', (s, i) => s.insertColumns(i, 30, 2)],
];

/**
 * Applies every edit in turn to every sheet and checks the values after each one. Returns the
 * edits that rebuilt the graph instead of following them in place.
 */
function checkEdits(workbook: Workbook, tolerance?: number): string[] {
	const session = createEditSession(workbook, { autoRowHeight: false });
	session.calculateNow({ full: true });
	const calc = session.calc;
	const invalidate = calc.invalidate.bind(calc);
	let rebuilt = false;
	calc.invalidate = () => {
		rebuilt = true;
		invalidate();
	};
	const rebuilds: string[] = [];
	for (let s = 0; s < workbook.sheets.length; s++)
		for (const [label, edit] of EDITS) {
			rebuilt = false;
			edit(session, s);
			if (rebuilt) rebuilds.push(`${label} on sheet ${s}`);
			expect(differences(workbook, tolerance), `${label} on sheet ${s}`).toEqual([]);
		}
	return rebuilds;
}

const fixture = (name: string): Promise<Workbook> =>
	loadXlsx(
		new Uint8Array(readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', name))),
	);

const put = (
	wb: Workbook,
	s: number,
	ref: string,
	content: CellValue | `=${string}`,
	extra: Partial<Cell> = {},
) => {
	const match = /^([A-Z]+)(\d+)$/.exec(ref);
	if (!match) throw new Error(ref);
	const col = [...(match[1] as string)].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
	const row = Number(match[2]) - 1;
	const cell: Cell =
		typeof content === 'string' && content.startsWith('=')
			? { value: null, formula: content.slice(1), ...extra }
			: { value: content, ...extra };
	putCell(wb.sheets[s] as never, row, col, cell);
};

/** Cross-sheet, whole-column, named, positional, intersection, lookup and array formulas. */
function synthetic(legacy: boolean): Workbook {
	const wb = createWorkbook({ sheets: ['Data', 'Report', 'Other Sheet'] });
	const f = legacy ? { legacyFormula: true as const } : {};
	for (let r = 1; r <= 20; r++) {
		put(wb, 0, `A${r}`, r * 3);
		put(wb, 0, `B${r}`, `=A${r}*2`, f);
		put(wb, 0, `C${r}`, r === 1 ? '=B1' : `=C${r - 1}+B${r}`, f);
		put(wb, 0, `D${r}`, `=$A$1+A${r}`, f);
	}
	const formulas = [
		'=SUM(A1:A20)',
		'=SUM(A:A)',
		'=INDEX(A:A,5)',
		'=ROW(A10)',
		'=ROW()',
		'=COLUMN(D1)',
		'=MATCH(21,A1:A20,0)',
		'=FORMULATEXT(B5)',
		"='Other Sheet'!A1+Report!B3",
		'=COUNTBLANK(A1:A25)',
		'=OFFSET(A1,4,0)',
		'=INDIRECT("A5")',
		'=SUM(Data)',
		'=CELL("address",A5)',
		legacy ? '=A1:A20' : '=@A1:A20',
		'=SUM(1:3)',
		'=VLOOKUP(12,A1:B20,2,FALSE)',
		'=SUM(Data:Report!A1)',
		'=SUM(A3 : A6)',
		'=ROWS(A2:A9)',
		'=SUM(INDEX(A:A,2):INDEX(A:A,8))',
		'=SHEET("Report")',
	];
	formulas.forEach((formula, i) => put(wb, 0, `F${i + 1}`, formula as `=${string}`, f));
	put(wb, 0, 'H1', '=A1:A3*2', {
		arrayRange: { start: { row: 0, col: 7 }, end: { row: 2, col: 7 } },
	});
	wb.definedNames.push({ name: 'Data', formula: 'Data!$A$1:$A$20' });
	for (let r = 1; r <= 8; r++) {
		put(wb, 1, `A${r}`, `=Data!A${r}+Data!C${r}`, f);
		put(wb, 1, `B${r}`, `=SUM(Data!A$1:A${r})`, f);
	}
	put(wb, 1, 'C1', '=SUM(Data!A:A)', f);
	put(wb, 1, 'C2', legacy ? '=Data!A1:A20' : '=@Data!A1:A20', f);
	put(wb, 1, 'C3', '=SUM(Data!F1:F8)', f);
	put(wb, 1, 'C4', '=COUNTIF(Data!B:B,">10")', f);
	put(wb, 2, 'A1', 7);
	put(wb, 2, 'B1', "='Report'!A3*2", f);
	return wb;
}

describe('structural edits recalculate in place', () => {
	for (const name of ['excel-features.xlsx', 'openpyxl-features.xlsx', 'excel-1904.xlsx'])
		it(`match a full recalculation on ${name}`, async () => {
			// Only undoing a row or column edit rebuilds.
			const rebuilds = checkEdits(await fixture(name));
			expect(rebuilds.filter((edit) => !edit.startsWith('undo '))).toEqual([]);
		});

	for (const legacy of [true, false])
		it(`match a full recalculation on a synthetic workbook (${legacy ? 'legacy' : 'dynamic'})`, () => {
			// Besides undoing row and column edits, only edits that cut the array formula H1:H3
			// (H2:H4 after the first insert) rebuild.
			const rebuilds = checkEdits(synthetic(legacy));
			expect(rebuilds.filter((edit) => !edit.startsWith('undo '))).toEqual([
				'insert rows inside on sheet 0',
				'delete a row on sheet 0',
				'redo on sheet 0',
				'delete rows inside on sheet 0',
			]);
		});

	it('match a full recalculation with spills, tables and iterated circular references', () => {
		const wb = synthetic(false);
		put(wb, 2, 'D1', '=SEQUENCE(4)');
		put(wb, 2, 'E1', '=SUM(D1#)');
		put(wb, 1, 'E1', '=E2*0.5+Data!A2');
		put(wb, 1, 'E2', '=E1*0.5+1');
		wb.iterate = { count: 100, delta: 1e-9 };
		const session = createEditSession(wb, { autoRowHeight: false });
		session.createTable(0, { start: { row: 0, col: 0 }, end: { row: 19, col: 1 } }, false);
		const table = wb.sheets[0]?.tables[0]?.name ?? '';
		put(wb, 2, 'F1', `=SUM(${table}[Column1])` as `=${string}`);
		// Edits on the sheet with the spill rebuild; the others still follow in place.
		const rebuilds = checkEdits(wb, 1e-6).filter((edit) => !edit.startsWith('undo '));
		expect(rebuilds.filter((edit) => !edit.endsWith('sheet 2'))).toEqual([
			'insert rows inside on sheet 0',
			'delete a row on sheet 0',
			'redo on sheet 0',
			'delete rows inside on sheet 0',
		]);
	});

	it('evaluates only what the edit can change', () => {
		const wb = synthetic(true);
		const session = createEditSession(wb, { autoRowHeight: false });
		session.calculateNow({ full: true });
		const data = wb.sheets[0];
		const report = wb.sheets[1];
		if (!data || !report) throw new Error('sheets');
		// Mark results the edit cannot change; a re-evaluation would overwrite the marks.
		const mark = (sheet: typeof data, row: number, col: number) => {
			const cell = getCell(sheet, row, col);
			if (cell) cell.value = -1;
		};
		mark(data, 15, 2); // C16, a running total below the edit: moves, same value
		mark(report, 0, 0); // Report!A1 reads Data!A1, above the edit
		session.insertRows(0, 10, 1);
		expect(getCell(data, 16, 2)?.value).toBe(-1);
		expect(getCell(report, 0, 0)?.value).toBe(-1);
		expect(getCell(data, 4, 5)?.value).toBe(5); // ROW() in F5 did not move
		expect(getCell(data, 0, 5)?.value).toBe(630); // SUM(A1:A21) took the new row
		expect(getCell(data, 3, 5)?.value).toBe(10); // ROW(A10) is above the edit
		session.calculateNow();
		expect(getCell(data, 16, 2)?.value).toBe(-1); // F9 recalculates only what changed
		session.calculateNow({ full: true });
		expect(getCell(data, 16, 2)?.value).not.toBe(-1);
	});
});
