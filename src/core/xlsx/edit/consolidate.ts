import { type CellAddress, type CellRange, MAX_COL, MAX_ROW, normalizeRange } from '../address';
import { getCell } from '../cells';
import type { CellValue, Workbook } from '../model';
import { type EditContext, displayText, ensureCell, pruneCell, sheetAt } from './context';
import { ErrorSignal, type Scalar, aggregate } from './deps';
import {
	type SubtotalFunctionCode,
	type SubtotalFunctionInput,
	resolveSubtotalFunction,
} from './subtotal-fn';

/** One source range fed into {@link consolidateRanges}. */
export interface ConsolidateSource {
	/** Zero-based worksheet index of this source range. */
	sheet: number;
	/** Rectangular range on that worksheet. */
	range: CellRange;
}

export interface ConsolidateOptions {
	/** Aggregation function name or SUBTOTAL code (1-11). */
	fn: SubtotalFunctionInput;
	/** Match columns by top-row labels (case-insensitive, default false). */
	useTopRow?: boolean;
	/** Match rows by left-column labels (case-insensitive, default false). */
	useLeftColumn?: boolean;
}

export interface ConsolidateResult {
	/** Destination range written. */
	range: CellRange;
	/** Number of rows written (including top header row when `useTopRow` is true). */
	rows: number;
	/** Number of columns written (including left label column when `useLeftColumn` is true). */
	cols: number;
}

interface NormalizedSource {
	sheet: number;
	range: CellRange;
}

function evaluateAggregate(code: SubtotalFunctionCode, values: Scalar[]): CellValue {
	if (!values.length) return null;
	try {
		return aggregate(code, values);
	} catch (error) {
		if (error instanceof ErrorSignal) return error.value;
		throw error;
	}
}

function readValue(workbook: Workbook, sheetIndex: number, row: number, col: number): Scalar {
	return getCell(sheetAt(workbook, sheetIndex), row, col)?.value ?? null;
}

function readLabel(workbook: Workbook, sheetIndex: number, row: number, col: number): string {
	return displayText(workbook, getCell(sheetAt(workbook, sheetIndex), row, col)).trim();
}

function collectByPosition(
	workbook: Workbook,
	sources: readonly NormalizedSource[],
	code: SubtotalFunctionCode,
): CellValue[][] {
	let outRows = 0;
	let outCols = 0;
	for (const { range } of sources) {
		outRows = Math.max(outRows, range.end.row - range.start.row + 1);
		outCols = Math.max(outCols, range.end.col - range.start.col + 1);
	}
	const grid: CellValue[][] = [];
	for (let r = 0; r < outRows; r++) {
		const rowVals: CellValue[] = [];
		for (let c = 0; c < outCols; c++) {
			const values: Scalar[] = [];
			for (const src of sources) {
				const sr = src.range.start.row + r;
				const sc = src.range.start.col + c;
				if (sr > src.range.end.row || sc > src.range.end.col) continue;
				const v = readValue(workbook, src.sheet, sr, sc);
				if (v !== null) values.push(v);
			}
			rowVals.push(evaluateAggregate(code, values));
		}
		grid.push(rowVals);
	}
	return grid;
}

function indexLabel(map: Map<string, number>, labels: string[], raw: string): number {
	const key = raw.toLowerCase();
	const existing = map.get(key);
	if (existing !== undefined) return existing;
	const idx = labels.length;
	map.set(key, idx);
	labels.push(raw);
	return idx;
}

function pushBucket(
	buckets: Map<number, Map<number, Scalar[]>>,
	r: number,
	c: number,
	value: Scalar,
): void {
	if (value === null) return;
	let rowMap = buckets.get(r);
	if (!rowMap) {
		rowMap = new Map();
		buckets.set(r, rowMap);
	}
	let list = rowMap.get(c);
	if (!list) {
		list = [];
		rowMap.set(c, list);
	}
	list.push(value);
}

function collectByCategory(
	workbook: Workbook,
	sources: readonly NormalizedSource[],
	code: SubtotalFunctionCode,
	useTopRow: boolean,
	useLeftColumn: boolean,
): CellValue[][] {
	const colMap = new Map<string, number>();
	const colLabels: string[] = [];
	const rowMap = new Map<string, number>();
	const rowLabels: string[] = [];
	const buckets = new Map<number, Map<number, Scalar[]>>();
	let maxDataRows = 0;
	let maxDataCols = 0;

	for (const src of sources) {
		const { range } = src;
		const dataRowStart = useTopRow ? range.start.row + 1 : range.start.row;
		const dataColStart = useLeftColumn ? range.start.col + 1 : range.start.col;
		const dataRows = Math.max(0, range.end.row - dataRowStart + 1);
		const dataCols = Math.max(0, range.end.col - dataColStart + 1);
		if (!useLeftColumn) maxDataRows = Math.max(maxDataRows, dataRows);
		if (!useTopRow) maxDataCols = Math.max(maxDataCols, dataCols);

		const colIndices: (number | undefined)[] = [];
		if (useTopRow) {
			for (let c = dataColStart; c <= range.end.col; c++) {
				const label = readLabel(workbook, src.sheet, range.start.row, c);
				colIndices.push(label ? indexLabel(colMap, colLabels, label) : undefined);
			}
		}

		for (let r = dataRowStart; r <= range.end.row; r++) {
			let outRow: number;
			if (useLeftColumn) {
				const label = readLabel(workbook, src.sheet, r, range.start.col);
				if (!label) continue;
				outRow = indexLabel(rowMap, rowLabels, label);
			} else {
				outRow = r - dataRowStart;
			}
			for (let c = dataColStart; c <= range.end.col; c++) {
				const outCol = useTopRow ? colIndices[c - dataColStart] : c - dataColStart;
				if (outCol === undefined) continue;
				pushBucket(buckets, outRow, outCol, readValue(workbook, src.sheet, r, c));
			}
		}
	}

	const numDataRows = useLeftColumn ? rowLabels.length : maxDataRows;
	const numDataCols = useTopRow ? colLabels.length : maxDataCols;
	const grid: CellValue[][] = [];

	if (useTopRow) {
		const header: CellValue[] = useLeftColumn ? [null, ...colLabels] : [...colLabels];
		if (header.length) grid.push(header);
	}
	for (let r = 0; r < numDataRows; r++) {
		const rowVals: CellValue[] = useLeftColumn ? [rowLabels[r] ?? null] : [];
		const rowBuckets = buckets.get(r);
		for (let c = 0; c < numDataCols; c++)
			rowVals.push(evaluateAggregate(code, rowBuckets?.get(c) ?? []));
		grid.push(rowVals);
	}
	return grid;
}

/**
 * Excel's Data > Consolidate: aggregates values across multiple source ranges (on the same or
 * different sheets) into `destination`, either by position or by matching top-row / left-column
 * category labels case-insensitively. One undo step.
 */
export function consolidateRanges(
	ctx: EditContext,
	destSheet: number,
	destination: CellAddress,
	sources: readonly ConsolidateSource[],
	options: ConsolidateOptions,
): ConsolidateResult {
	const target = sheetAt(ctx.workbook, destSheet);
	if (
		!Number.isInteger(destination.row) ||
		!Number.isInteger(destination.col) ||
		destination.row < 0 ||
		destination.row > MAX_ROW ||
		destination.col < 0 ||
		destination.col > MAX_COL
	)
		throw new RangeError('Consolidate destination must be a valid cell address.');
	if (!sources.length) throw new RangeError('Consolidate requires at least one source range.');

	const meta = resolveSubtotalFunction(options.fn);
	const normalized: NormalizedSource[] = sources.map((src) => {
		sheetAt(ctx.workbook, src.sheet);
		return { sheet: src.sheet, range: normalizeRange(src.range) };
	});

	const useTopRow = options.useTopRow ?? false;
	const useLeftColumn = options.useLeftColumn ?? false;
	const grid =
		useTopRow || useLeftColumn
			? collectByCategory(ctx.workbook, normalized, meta.code, useTopRow, useLeftColumn)
			: collectByPosition(ctx.workbook, normalized, meta.code);

	const rows = grid.length;
	const cols = grid[0]?.length ?? 0;
	if (!rows || !cols)
		return { range: { start: { ...destination }, end: { ...destination } }, rows: 0, cols: 0 };

	const destRange: CellRange = {
		start: { row: destination.row, col: destination.col },
		end: { row: destination.row + rows - 1, col: destination.col + cols - 1 },
	};
	if (destRange.end.row > MAX_ROW || destRange.end.col > MAX_COL)
		throw new RangeError('Consolidated output exceeds the worksheet grid.');

	ctx.run(
		'Consolidate',
		'cells',
		[{ kind: 'cells', sheet: destSheet, ranges: [destRange] }],
		() => {
			for (let r = 0; r < rows; r++) {
				const rowVals = grid[r];
				if (!rowVals) continue;
				for (let c = 0; c < cols; c++) {
					const val = rowVals[c] ?? null;
					const row = destination.row + r;
					const col = destination.col + c;
					if (val === null) {
						const existing = getCell(target, row, col);
						if (existing) {
							existing.value = null;
							delete existing.formula;
							delete existing.richText;
							pruneCell(target, row, col);
						}
					} else {
						const cell = ensureCell(target, row, col);
						cell.value = val;
						delete cell.formula;
						delete cell.richText;
					}
				}
			}
		},
		{ sheet: destSheet, ranges: [destRange] },
	);

	return { range: destRange, rows, cols };
}
