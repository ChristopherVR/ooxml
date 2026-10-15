// What-if data tables (Data > What-If Analysis > Data Table): `{=TABLE(row input, column input)}`.
// The result cells hold plain values that are recomputed here: for each one the input cells are
// given the header values, the workbook recalculates, and the formula cell is read; the real
// inputs are restored afterwards.
import {
	type CellAddress,
	type CellRange,
	formatAddress,
	normalizeRange,
	parseAddress,
	rangeContains,
	rangesIntersect,
} from '../address';
import { deleteCell, getCell } from '../cells';
import type { Cell, CellValue, DataTable, Workbook, Worksheet } from '../model';
import { type EditContext, ensureCell, sheetAt } from './context';
import type { CalcEngine, CellPosition } from './deps';
import { cellRange } from './range-math';

export const DATA_TABLE_PART_ERROR = 'Cannot change part of a data table.';

/** The data table whose result range holds a cell. */
export function dataTableAt(sheet: Worksheet, row: number, col: number): DataTable | undefined {
	return sheet.dataTables?.find((t) => rangeContains(t.ref, { row, col }));
}

/** The formula Excel shows for a data table's result cells, without braces: `TABLE(B1,A1)`. */
export function dataTableFormula(table: DataTable): string {
	if (table.dt2D) return `TABLE(${table.r1},${table.r2 ?? ''})`;
	return table.dtr ? `TABLE(${table.r1},)` : `TABLE(,${table.r1})`;
}

/** What the formula bar shows for a cell inside a data table, or undefined outside one. */
export function dataTableFormulaText(
	sheet: Worksheet,
	row: number,
	col: number,
): string | undefined {
	const table = dataTableAt(sheet, row, col);
	return table ? `{=${dataTableFormula(table)}}` : undefined;
}

/** Whether an edit confined to `range` would change part of a data table (not all of it). */
export function cutsDataTable(sheet: Worksheet, range: CellRange): boolean {
	const r = normalizeRange(range);
	return (sheet.dataTables ?? []).some(
		(t) =>
			rangesIntersect(t.ref, r) && !(rangeContains(r, t.ref.start) && rangeContains(r, t.ref.end)),
	);
}

/** Reads an input cell reference (`B1`, `$B$1`, `Sheet1!$B$1`) of the same sheet as `A1` text. */
export function parseInputCell(sheet: Worksheet, text: string): CellAddress | undefined {
	let ref = text.trim().replace(/^=/, '');
	const bang = ref.lastIndexOf('!');
	if (bang >= 0) {
		const name = ref
			.slice(0, bang)
			.replace(/^'(.*)'$/, '$1')
			.replace(/''/g, "'");
		if (name.toLowerCase() !== sheet.name.toLowerCase()) return undefined;
		ref = ref.slice(bang + 1);
	}
	return parseAddress(ref.replace(/\$/g, ''));
}

interface Saved {
	row: number;
	col: number;
	existed: boolean;
	cell: Cell | undefined;
	fields: Pick<Cell, 'value' | 'formula' | 'arrayRange' | 'legacyFormula' | 'dynamicArray'>;
}

function save(sheet: Worksheet, row: number, col: number): Saved {
	const cell = getCell(sheet, row, col);
	return {
		row,
		col,
		existed: !!cell,
		cell,
		fields: cell
			? {
					value: cell.value,
					...(cell.formula !== undefined ? { formula: cell.formula } : {}),
					...(cell.arrayRange ? { arrayRange: cell.arrayRange } : {}),
					...(cell.legacyFormula ? { legacyFormula: true as const } : {}),
					...(cell.dynamicArray ? { dynamicArray: true as const } : {}),
				}
			: { value: null },
	};
}

function restore(sheet: Worksheet, saved: Saved): void {
	if (!saved.existed) {
		deleteCell(sheet, saved.row, saved.col);
		return;
	}
	const cell = getCell(sheet, saved.row, saved.col);
	if (!cell) return;
	delete cell.formula;
	delete cell.arrayRange;
	delete cell.legacyFormula;
	delete cell.dynamicArray;
	Object.assign(cell, saved.fields);
}

const valueAt = (sheet: Worksheet, row: number, col: number): CellValue =>
	getCell(sheet, row, col)?.value ?? null;

/**
 * Recomputes one data table's result cells in place (no undo step) and returns the positions
 * whose value changed.
 */
export function computeDataTable(
	workbook: Workbook,
	calc: CalcEngine,
	sheetIndex: number,
	table: DataTable,
): CellPosition[] {
	const sheet = sheetAt(workbook, sheetIndex);
	const { ref } = table;
	const first = parseInputCell(sheet, table.r1);
	const second = table.dt2D && table.r2 ? parseInputCell(sheet, table.r2) : undefined;
	if (!first || (table.dt2D && !second) || table.del1 || table.del2) return [];
	const inputs = [first, ...(second ? [second] : [])];
	const saved = inputs.map((p) => save(sheet, p.row, p.col));
	const top = ref.start.row - 1;
	const left = ref.start.col - 1;
	const changed: CellPosition[] = [];
	const positions = inputs.map((p) => ({ sheet: sheetIndex, row: p.row, col: p.col }));
	const set = (address: CellAddress, value: CellValue): void => {
		const cell = ensureCell(sheet, address.row, address.col);
		delete cell.formula;
		delete cell.arrayRange;
		delete cell.legacyFormula;
		delete cell.dynamicArray;
		cell.value = value;
	};
	try {
		for (let row = ref.start.row; row <= ref.end.row; row++) {
			for (let col = ref.start.col; col <= ref.end.col; col++) {
				let formulaAt: CellAddress;
				if (table.dt2D) {
					set(first, valueAt(sheet, top, col));
					if (second) set(second, valueAt(sheet, row, left));
					formulaAt = { row: top, col: left };
				} else if (table.dtr) {
					set(first, valueAt(sheet, top, col));
					formulaAt = { row, col: left };
				} else {
					set(first, valueAt(sheet, row, left));
					formulaAt = { row: top, col };
				}
				calc.recalculateFrom(positions);
				const value = valueAt(sheet, formulaAt.row, formulaAt.col);
				const target = ensureCell(sheet, row, col);
				if (target.value !== value) {
					target.value = value;
					changed.push({ sheet: sheetIndex, row, col });
				}
			}
		}
	} finally {
		for (const s of saved) restore(sheet, s);
		calc.recalculateFrom(positions);
	}
	return changed;
}

/** Recomputes every data table of the workbook and lets dependents of the results follow. */
export function refreshDataTables(workbook: Workbook, calc: CalcEngine): void {
	const changed: CellPosition[] = [];
	workbook.sheets.forEach((sheet, index) => {
		for (const table of sheet.dataTables ?? [])
			changed.push(...computeDataTable(workbook, calc, index, table));
	});
	if (changed.length) calc.recalculateFrom(changed);
}

export interface DataTableRequest {
	/** The whole selection: header row and column included. */
	range: CellRange;
	rowInput?: string;
	colInput?: string;
}

export type DataTableProblem = 'range-too-small' | 'no-input' | 'bad-input' | 'input-inside';

const MESSAGES: Record<DataTableProblem, string> = {
	'range-too-small': 'Select a range with input values and a formula around the results.',
	'no-input': 'Enter a row input cell, a column input cell, or both.',
	'bad-input': 'Input cell reference is not valid.',
	'input-inside': 'An input cell cannot be inside the data table.',
};

export const dataTableMessage = (problem: DataTableProblem): string => MESSAGES[problem];

function resolve(
	sheet: Worksheet,
	request: DataTableRequest,
): { problem: DataTableProblem } | { table: DataTable } {
	const range = normalizeRange(request.range);
	if (range.end.row - range.start.row < 1 || range.end.col - range.start.col < 1)
		return { problem: 'range-too-small' };
	const rowText = request.rowInput?.trim() ?? '';
	const colText = request.colInput?.trim() ?? '';
	if (!rowText && !colText) return { problem: 'no-input' };
	const row = rowText ? parseInputCell(sheet, rowText) : undefined;
	const col = colText ? parseInputCell(sheet, colText) : undefined;
	if ((rowText && !row) || (colText && !col)) return { problem: 'bad-input' };
	const ref = cellRange(range.start.row + 1, range.start.col + 1, range.end.row, range.end.col);
	if ([row, col].some((p) => p && rangeContains(ref, p))) return { problem: 'input-inside' };
	const name = (p: CellAddress): string => formatAddress(p);
	const table: DataTable = { ref, r1: '' };
	if (row && col) {
		table.dt2D = true;
		table.dtr = true;
		table.r1 = name(row);
		table.r2 = name(col);
	} else if (row) {
		table.dtr = true;
		table.r1 = name(row);
	} else if (col) table.r1 = name(col);
	return { table };
}

/** Why a request cannot create a data table, or null. */
export function checkDataTable(workbook: Workbook, sheetIndex: number, request: DataTableRequest) {
	const result = resolve(sheetAt(workbook, sheetIndex), request);
	return 'problem' in result ? result.problem : null;
}

/**
 * Data > What-If Analysis > Data Table: turns the selection into `{=TABLE(row, column)}` and
 * computes it. An existing table inside the selection is replaced. One undo step.
 */
export function createDataTable(ctx: EditContext, s: number, request: DataTableRequest): void {
	const sheet = sheetAt(ctx.workbook, s);
	const result = resolve(sheet, request);
	if ('problem' in result) throw new Error(MESSAGES[result.problem]);
	const { table } = result;
	const range = normalizeRange(request.range);
	ctx.run(
		'Data Table',
		'cells',
		[
			{ kind: 'cells', sheet: s, ranges: [range] },
			{ kind: 'parts', sheet: s, parts: ['dataTables'] },
		],
		() => {
			sheet.dataTables = (sheet.dataTables ?? []).filter((t) => !rangesIntersect(t.ref, range));
			sheet.dataTables.push(table);
			for (let row = table.ref.start.row; row <= table.ref.end.row; row++)
				for (let col = table.ref.start.col; col <= table.ref.end.col; col++) {
					const cell = ensureCell(sheet, row, col);
					delete cell.formula;
					delete cell.arrayRange;
					delete cell.legacyFormula;
					delete cell.dynamicArray;
				}
			computeDataTable(ctx.workbook, ctx.calc, s, table);
		},
		{ sheet: s, ranges: [range] },
	);
}
