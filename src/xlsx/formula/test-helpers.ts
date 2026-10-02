// Shared fixtures for the formula tests (not part of the public API).
import { parseAddress, parseRange } from '../address.js';
import { getCell, putCell } from '../cells.js';
import type { CellValue, Workbook } from '../model.js';
import { createWorkbook } from '../workbook.js';
import { type CalcEngine, createCalcEngine } from './engine.js';
import { err } from './values.js';

export type CellSpec = CellValue | `=${string}`;

/** A workbook from `{ 'A1': 1, 'Sheet2!B2': '=A1*2' }`; strings starting with `=` are formulas. */
export function book(
	cells: Record<string, CellSpec> = {},
	sheets: string[] = ['Sheet1', 'Sheet2'],
): Workbook {
	const wb = createWorkbook({ sheets });
	for (const [key, value] of Object.entries(cells)) set(wb, key, value);
	return wb;
}

export function locate(wb: Workbook, key: string): { sheet: number; row: number; col: number } {
	const bang = key.lastIndexOf('!');
	const sheetName = bang >= 0 ? key.slice(0, bang) : (wb.sheets[0]?.name ?? 'Sheet1');
	const sheet = wb.sheets.findIndex((s) => s.name === sheetName);
	const address = parseAddress(bang >= 0 ? key.slice(bang + 1) : key);
	if (sheet < 0 || !address) throw new Error(`bad cell key ${key}`);
	return { sheet, ...address };
}

export function set(wb: Workbook, key: string, value: CellSpec): void {
	const { sheet, row, col } = locate(wb, key);
	const ws = wb.sheets[sheet];
	if (!ws) throw new Error('no sheet');
	if (typeof value === 'string' && value.startsWith('='))
		putCell(ws, row, col, { value: null, formula: value.slice(1) });
	else putCell(ws, row, col, { value });
}

export function get(wb: Workbook, key: string): CellValue {
	const { sheet, row, col } = locate(wb, key);
	const ws = wb.sheets[sheet];
	return ws ? (getCell(ws, row, col)?.value ?? null) : null;
}

/** Values of a range as rows. */
export function grid(wb: Workbook, ref: string, sheetIndex = 0): CellValue[][] {
	const range = parseRange(ref);
	const ws = wb.sheets[sheetIndex];
	if (!range || !ws) throw new Error('bad range');
	const out: CellValue[][] = [];
	for (let r = range.start.row; r <= range.end.row; r++) {
		const line: CellValue[] = [];
		for (let c = range.start.col; c <= range.end.col; c++)
			line.push(getCell(ws, r, c)?.value ?? null);
		out.push(line);
	}
	return out;
}

/** A deterministic engine (fixed clock 2026-10-03 12:00, seeded random). */
export function engine(wb: Workbook): CalcEngine {
	let seed = 42;
	return createCalcEngine(wb, {
		now: () => new Date(2026, 9, 3, 12, 0, 0),
		random: () => {
			seed = (seed * 16807) % 2147483647;
			return (seed - 1) / 2147483646;
		},
	});
}

/** Puts `formula` in Sheet1!AA1 of a workbook with `cells`, recalculates, and returns its value. */
export function calc(formula: string, cells: Record<string, CellSpec> = {}): CellValue {
	const wb = book(cells);
	set(wb, 'AA1', `=${formula}`);
	engine(wb).recalculateAll();
	return get(wb, 'AA1');
}

/** Like {@link calc} but returns the spilled block (`rows` x `cols`) anchored at AA1. */
export function calcArray(
	formula: string,
	cells: Record<string, CellSpec> = {},
	rows = 1,
	cols = 1,
): CellValue[][] {
	const wb = book(cells);
	set(wb, 'AA1', `=${formula}`);
	engine(wb).recalculateAll();
	const ws = wb.sheets[0];
	const out: CellValue[][] = [];
	for (let r = 0; r < rows; r++) {
		const line: CellValue[] = [];
		for (let c = 0; c < cols; c++) line.push(ws ? (getCell(ws, r, 26 + c)?.value ?? null) : null);
		out.push(line);
	}
	return out;
}

export const E = {
	DIV0: err('#DIV/0!'),
	VALUE: err('#VALUE!'),
	REF: err('#REF!'),
	NAME: err('#NAME?'),
	NUM: err('#NUM!'),
	NA: err('#N/A'),
	NULL: err('#NULL!'),
	SPILL: err('#SPILL!'),
	CALC: err('#CALC!'),
};
