import { type CellAddress, type CellRange, MAX_COL, MAX_ROW, normalizeRange } from '../address';
import { getCell } from '../cells';
import type { Workbook, Worksheet } from '../model';
import { syncTableHeader, writeScopes } from './cell-values';
import { type EditContext, displayText, ensureCell, sheetAt } from './context';
import {
	type FlashFillExample,
	type FlashFillProgram,
	inferFlashFillProgram,
} from './flash-fill-infer';

export interface FlashFillOptions {
	/** Explicit zero-based source column indices (defaults to contiguous adjacent columns). */
	sourceColumns?: readonly number[];
	/** Treat the first row of the region as a header row and skip it. */
	hasHeader?: boolean;
}

export interface FlashFillResult {
	/** Number of cells filled (0 when no consistent pattern was found or nothing was empty). */
	filled: number;
	/** Range covering the filled cells, or undefined when `filled === 0`. */
	range: CellRange | undefined;
}

function isNonEmpty(sheet: Worksheet, row: number, col: number): boolean {
	const cell = getCell(sheet, row, col);
	return !!cell && cell.value !== null && cell.value !== '';
}

function detectSourceColumns(sheet: Worksheet, targetCol: number, anchorRow: number): number[] {
	const hasNear = (col: number): boolean => {
		for (let r = Math.max(0, anchorRow - 1); r <= Math.min(MAX_ROW, anchorRow + 1); r++)
			if (isNonEmpty(sheet, r, col)) return true;
		return false;
	};
	const left: number[] = [];
	for (let c = targetCol - 1; c >= 0 && hasNear(c); c--) left.push(c);
	left.reverse();
	const right: number[] = [];
	for (let c = targetCol + 1; c <= MAX_COL && hasNear(c); c++) right.push(c);
	return [...left, ...right];
}

function detectRowSpan(
	sheet: Worksheet,
	anchorRow: number,
	targetCol: number,
	sourceCols: readonly number[],
): { startRow: number; endRow: number } {
	const rowHasData = (row: number): boolean => {
		if (isNonEmpty(sheet, row, targetCol)) return true;
		for (const c of sourceCols) if (isNonEmpty(sheet, row, c)) return true;
		return false;
	};
	let startRow = anchorRow;
	let endRow = anchorRow;
	while (startRow > 0 && rowHasData(startRow - 1)) startRow--;
	while (endRow < MAX_ROW && rowHasData(endRow + 1)) endRow++;
	return { startRow, endRow };
}

function readInputs(
	workbook: Workbook,
	sheet: Worksheet,
	row: number,
	sourceCols: readonly number[],
): string[] {
	return sourceCols.map((col) => displayText(workbook, getCell(sheet, row, col)));
}

function tryInfer(
	workbook: Workbook,
	sheet: Worksheet,
	startRow: number,
	endRow: number,
	targetCol: number,
	sourceCols: readonly number[],
	hasHeader: boolean | undefined,
): { program: FlashFillProgram; firstDataRow: number } | undefined {
	const collect = (from: number): FlashFillExample[] => {
		const list: FlashFillExample[] = [];
		for (let row = from; row <= endRow; row++) {
			const output = displayText(workbook, getCell(sheet, row, targetCol));
			if (!output) continue;
			const inputs = readInputs(workbook, sheet, row, sourceCols);
			if (inputs.some((v) => v.length > 0)) list.push({ inputs, output });
		}
		return list;
	};

	if (hasHeader === true) {
		const examples = collect(startRow + 1);
		const program = inferFlashFillProgram(examples);
		return program ? { program, firstDataRow: startRow + 1 } : undefined;
	}

	const allExamples = collect(startRow);
	const program = inferFlashFillProgram(allExamples);
	if (program) return { program, firstDataRow: startRow };

	// Auto header detection: if the first row has a label in the target column that does not fit
	// the pattern of subsequent examples, retry skipping the first row as a header.
	if (hasHeader === undefined && startRow < endRow && isNonEmpty(sheet, startRow, targetCol)) {
		const withoutHeader = collect(startRow + 1);
		if (withoutHeader.length > 0) {
			const fallback = inferFlashFillProgram(withoutHeader);
			if (fallback) return { program: fallback, firstDataRow: startRow + 1 };
		}
	}
	return undefined;
}

/**
 * Excel's Data > Flash Fill (Ctrl+E): infers a string transformation from non-empty example cells
 * in the target column and adjacent contiguous source column(s), then fills empty cells in the
 * contiguous data region. Returns `{ filled: 0, range: undefined }` as a no-op when no consistent
 * pattern is found. One undo step.
 */
export function flashFill(
	ctx: EditContext,
	s: number,
	target: CellAddress | CellRange,
	options: FlashFillOptions = {},
): FlashFillResult {
	const sheet = sheetAt(ctx.workbook, s);
	const isRange = 'start' in target;
	const norm = isRange ? normalizeRange(target) : { start: target, end: target };
	if (norm.start.col !== norm.end.col)
		throw new RangeError('Flash Fill target must be a single cell or single-column range.');
	const targetCol = norm.start.col;
	if (!Number.isInteger(targetCol) || targetCol < 0 || targetCol > MAX_COL)
		throw new RangeError('Flash Fill target column is out of range.');

	const sourceCols = options.sourceColumns
		? [...new Set(options.sourceColumns)].filter((c) => c !== targetCol && c >= 0 && c <= MAX_COL)
		: detectSourceColumns(sheet, targetCol, norm.start.row);
	if (!sourceCols.length) return { filled: 0, range: undefined };

	const singleCell = !isRange || norm.start.row === norm.end.row;
	const { startRow, endRow } = singleCell
		? detectRowSpan(sheet, norm.start.row, targetCol, sourceCols)
		: { startRow: norm.start.row, endRow: norm.end.row };

	const inferred = tryInfer(
		ctx.workbook,
		sheet,
		startRow,
		endRow,
		targetCol,
		sourceCols,
		options.hasHeader,
	);
	if (!inferred) return { filled: 0, range: undefined };

	const writes: { row: number; value: string }[] = [];
	for (let row = inferred.firstDataRow; row <= endRow; row++) {
		if (isNonEmpty(sheet, row, targetCol)) continue;
		const inputs = readInputs(ctx.workbook, sheet, row, sourceCols);
		if (!inputs.some((v) => v.length > 0)) continue;
		const predicted = inferred.program(inputs);
		if (predicted !== undefined) writes.push({ row, value: predicted });
	}
	if (!writes.length) return { filled: 0, range: undefined };

	const fillRange: CellRange = {
		start: { row: writes[0]!.row, col: targetCol },
		end: { row: writes[writes.length - 1]!.row, col: targetCol },
	};

	ctx.run(
		'Flash Fill',
		'cells',
		writeScopes(sheet, s, fillRange),
		() => {
			for (const { row, value } of writes) {
				const cell = ensureCell(sheet, row, targetCol);
				cell.value = value;
				delete cell.formula;
				delete cell.richText;
				syncTableHeader(ctx, sheet, row, targetCol);
			}
		},
		{ sheet: s, ranges: [fillRange] },
	);

	return { filled: writes.length, range: fillRange };
}
