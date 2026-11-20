// The calculation on open run elsewhere (a worker) and brought back as only what it wrote: the
// values of formula cells, spilled cells and array ranges, so the page can keep the workbook it
// read itself instead of receiving a copy of the whole model.
import { cellKey, keyToAddress } from '../address';
import { deleteCell, getCell, putCell } from '../cells';
import type { CellValue, Workbook } from '../model';
import { calculateOnOpen, type OpenCalculation } from './calc-mode';
import { isSpilledCell, type SpilledCell } from './deps';

/** What a calculation wrote to one sheet, in flat arrays (cheap to post between threads). */
export interface CalculatedSheet {
	/** Row and column of each written cell, in pairs. */
	cells: Int32Array;
	/** The value of each written cell. */
	values: CellValue[];
	/** The spill anchor's row and column of each written cell, in pairs; -1 for none. */
	anchors: Int32Array;
	/** Row and column of each cell a stored spill left, in pairs: its value is cleared. */
	released: Int32Array;
}

/** What `calculateOnOpen` wrote to a workbook, to replay on another copy of the same file. */
export interface CalculatedCells {
	sheets: CalculatedSheet[];
	warnings: string[];
}

const NO_ANCHOR = { row: -1, col: -1 };

/** The cells a calculation may write on one sheet: formulas, spills and array formula ranges. */
function writtenOn(workbook: Workbook, s: number, before: ReadonlySet<number>): CalculatedSheet {
	const sheet = workbook.sheets[s];
	const cells: number[] = [];
	const anchors: number[] = [];
	const values: CellValue[] = [];
	const released: number[] = [];
	if (!sheet)
		return {
			cells: new Int32Array(),
			values,
			anchors: new Int32Array(),
			released: new Int32Array(),
		};
	const arrays: { row: number; col: number; endRow: number; endCol: number }[] = [];
	for (const [row, line] of sheet.rows)
		for (const [col, cell] of line) {
			const spilled = isSpilledCell(cell);
			if (!cell.formula && !spilled) continue;
			const anchor = spilled ? cell.spillAnchor : NO_ANCHOR;
			cells.push(row, col);
			anchors.push(anchor.row, anchor.col);
			values.push(cell.value);
			const range = cell.arrayRange;
			if (range && (range.end.row > range.start.row || range.end.col > range.start.col))
				arrays.push({ row, col, endRow: range.end.row, endCol: range.end.col });
		}
	for (const { row, col, endRow, endCol } of arrays)
		for (let r = row; r <= endRow; r++)
			for (let c = col; c <= endCol; c++) {
				const cell = getCell(sheet, r, c);
				if (!cell || cell.formula || isSpilledCell(cell)) continue;
				cells.push(r, c);
				anchors.push(-1, -1);
				values.push(cell.value);
			}
	for (const key of before) {
		const { row, col } = keyToAddress(key);
		if (!isSpilledCell(getCell(sheet, row, col))) released.push(row, col);
	}
	return {
		cells: Int32Array.from(cells),
		values,
		anchors: Int32Array.from(anchors),
		released: Int32Array.from(released),
	};
}

/**
 * Runs `calculateOnOpen` and returns, with what it did, the cells it wrote (when it ran to the
 * end), for `applyCalculatedCells` on another copy of the workbook read from the same file.
 */
export function calculateOnOpenCells(
	workbook: Workbook,
	options: { timeLimitMs?: number } = {},
): { calculation: OpenCalculation; cells?: CalculatedCells } {
	const spilledBefore = workbook.sheets.map((sheet) => {
		const keys = new Set<number>();
		for (const [row, line] of sheet.rows)
			for (const [col, cell] of line) if (isSpilledCell(cell)) keys.add(cellKey(row, col));
		return keys;
	});
	const calculation = calculateOnOpen(workbook, options);
	if (!calculation.ran || calculation.stopped) return { calculation };
	return {
		calculation,
		cells: {
			sheets: workbook.sheets.map((_, s) => writtenOn(workbook, s, spilledBefore[s] ?? new Set())),
			warnings: [...workbook.warnings],
		},
	};
}

/** Writes what a calculation on open wrote elsewhere into this copy of the workbook. */
export function applyCalculatedCells(workbook: Workbook, calculated: CalculatedCells): void {
	calculated.sheets.forEach((written, s) => {
		const sheet = workbook.sheets[s];
		if (!sheet) return;
		const { cells, values, anchors, released } = written;
		for (let i = 0; i < released.length; i += 2) {
			const row = released[i] as number;
			const col = released[i + 1] as number;
			const cell = getCell(sheet, row, col);
			if (!cell) continue;
			delete (cell as { spillAnchor?: unknown }).spillAnchor;
			cell.value = null;
			if (!cell.styleId && !cell.formula && !cell.richText && !cell.arrayRange)
				deleteCell(sheet, row, col);
		}
		for (let i = 0, v = 0; i < cells.length; i += 2, v++) {
			const row = cells[i] as number;
			const col = cells[i + 1] as number;
			const value = values[v] ?? null;
			const anchorRow = anchors[i] as number;
			const anchor = anchorRow < 0 ? undefined : { row: anchorRow, col: anchors[i + 1] as number };
			const cell = getCell(sheet, row, col);
			if (!cell) {
				const created: SpilledCell | { value: CellValue } = anchor
					? { value, spillAnchor: anchor }
					: { value };
				putCell(sheet, row, col, created);
				continue;
			}
			cell.value = value;
			if (anchor) (cell as { spillAnchor?: unknown }).spillAnchor = anchor;
			else if (isSpilledCell(cell)) delete (cell as { spillAnchor?: unknown }).spillAnchor;
		}
	});
	workbook.warnings = [...calculated.warnings];
}
