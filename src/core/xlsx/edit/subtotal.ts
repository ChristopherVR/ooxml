import { type CellRange, formatRange, normalizeRange, rangesIntersect } from '../address';
import { getCell, putCell } from '../cells';
import type { RowInfo, Worksheet } from '../model';
import { type EditContext, displayText, ensureCell, sheetAt } from './context';
import { MAX_OUTLINE_LEVEL } from './outline';
import { setPageBreak } from './page-breaks';
import { applyAxisShift } from './structure';
import {
	type SubtotalFunctionCode,
	type SubtotalFunctionInput,
	resolveSubtotalFunction,
} from './subtotal-fn';

export interface SubtotalOptions {
	/** Zero-based column index to group by ("At each change in"). */
	atEachChangeIn: number;
	/** Aggregation function name or SUBTOTAL code (1-11). */
	fn: SubtotalFunctionInput;
	/** Zero-based column indices that receive `=SUBTOTAL(...)` formulas. */
	addSubtotalTo: readonly number[];
	/** Remove existing subtotal rows before inserting new ones (default true). */
	replaceExisting?: boolean;
	/** Place summary rows below each group and Grand Total at the bottom (default true). */
	summaryBelowData?: boolean;
	/** Whether the first row of `range` is a header row (default true). */
	hasHeader?: boolean;
	/** Insert a horizontal page break between groups (default false). */
	pageBreakBetweenGroups?: boolean;
}

export interface SubtotalResult {
	/** Final range covering the original range plus any inserted summary rows. */
	range: CellRange;
	/** Zero-based row indices of the inserted group subtotal rows. */
	subtotalRows: number[];
	/** Zero-based row index of the Grand Total row, or undefined when no data rows exist. */
	grandTotalRow: number | undefined;
}

export interface RemoveSubtotalsResult {
	/** Number of subtotal rows removed. */
	removed: number;
	/** Resulting range after removing subtotal rows. */
	range: CellRange;
}

const SUBTOTAL_FORMULA_RE = /^\s*(?:_xlfn\.)?SUBTOTAL\s*\(/i;

function isSubtotalRow(sheet: Worksheet, row: number, colStart: number, colEnd: number): boolean {
	const cells = sheet.rows.get(row);
	if (!cells) return false;
	for (let col = colStart; col <= colEnd; col++) {
		const formula = cells.get(col)?.formula;
		if (formula !== undefined && SUBTOTAL_FORMULA_RE.test(formula)) return true;
	}
	return false;
}

function setRowOutline(sheet: Worksheet, row: number, level: number): void {
	const existing = sheet.rowInfo.get(row);
	if (!level) {
		if (!existing) return;
		const next: RowInfo = { ...existing };
		delete next.outlineLevel;
		delete next.collapsed;
		if (Object.keys(next).length) sheet.rowInfo.set(row, next);
		else sheet.rowInfo.delete(row);
		return;
	}
	const clamped = Math.min(MAX_OUTLINE_LEVEL, Math.max(1, level));
	sheet.rowInfo.set(row, { ...existing, outlineLevel: clamped });
}

function stripSubtotalsInPlace(
	ctx: EditContext,
	sheet: Worksheet,
	r: CellRange,
	firstData: number,
): number {
	const doomed: number[] = [];
	for (let row = firstData; row <= r.end.row; row++)
		if (isSubtotalRow(sheet, row, r.start.col, r.end.col)) doomed.push(row);
	for (let i = doomed.length - 1; i >= 0; i--) {
		const row = doomed[i];
		if (row !== undefined) applyAxisShift(ctx, sheet, { axis: 'row', at: row, count: -1 });
	}
	const newEnd = r.end.row - doomed.length;
	for (let row = r.start.row; row <= newEnd; row++) setRowOutline(sheet, row, 0);
	return doomed.length;
}

interface DataGroup {
	start: number;
	end: number;
	label: string;
}

interface PlacedGroup {
	subRow: number;
	dataStart: number;
	dataEnd: number;
	label: string;
}

function collectGroups(
	ctx: EditContext,
	sheet: Worksheet,
	firstData: number,
	lastRow: number,
	r: CellRange,
	groupCol: number,
): DataGroup[] {
	const groups: DataGroup[] = [];
	let current: DataGroup | undefined;
	let currentKey = '';
	for (let row = firstData; row <= lastRow; row++) {
		if (isSubtotalRow(sheet, row, r.start.col, r.end.col)) {
			current = undefined;
			continue;
		}
		const text = displayText(ctx.workbook, getCell(sheet, row, groupCol));
		const key = text.toLowerCase();
		if (current && key === currentKey) current.end = row;
		else {
			current = { start: row, end: row, label: text };
			currentKey = key;
			groups.push(current);
		}
	}
	return groups;
}

function insertSummaryRows(
	ctx: EditContext,
	sheet: Worksheet,
	groups: readonly DataGroup[],
	firstData: number,
	endRow: number,
	summaryBelow: boolean,
): { placed: PlacedGroup[]; grandTotalRow: number; finalEndRow: number } {
	const placed: PlacedGroup[] = [];
	let shift = 0;
	let grandTotalRow: number;
	if (summaryBelow) {
		for (const g of groups) {
			const dataStart = g.start + shift;
			const dataEnd = g.end + shift;
			const subRow = dataEnd + 1;
			applyAxisShift(ctx, sheet, { axis: 'row', at: subRow, count: 1 });
			placed.push({ subRow, dataStart, dataEnd, label: g.label });
			shift += 1;
		}
		grandTotalRow = endRow + shift + 1;
		applyAxisShift(ctx, sheet, { axis: 'row', at: grandTotalRow, count: 1 });
		shift += 1;
	} else {
		grandTotalRow = firstData;
		applyAxisShift(ctx, sheet, { axis: 'row', at: grandTotalRow, count: 1 });
		shift += 1;
		for (const g of groups) {
			const subRow = g.start + shift;
			applyAxisShift(ctx, sheet, { axis: 'row', at: subRow, count: 1 });
			shift += 1;
			placed.push({ subRow, dataStart: g.start + shift, dataEnd: g.end + shift, label: g.label });
		}
	}
	return { placed, grandTotalRow, finalEndRow: endRow + shift };
}

function writeSummaryRow(
	sheet: Worksheet,
	row: number,
	labelCol: number,
	label: string,
	code: SubtotalFunctionCode,
	valueCols: readonly number[],
	spanStart: number,
	spanEnd: number,
): void {
	const labelCell = ensureCell(sheet, row, labelCol);
	labelCell.value = label;
	delete labelCell.formula;
	delete labelCell.richText;
	for (const col of valueCols) {
		const ref = formatRange({ start: { row: spanStart, col }, end: { row: spanEnd, col } });
		const existing = getCell(sheet, row, col);
		const cell = existing ? { ...existing } : { value: null };
		cell.value = null;
		cell.formula = `SUBTOTAL(${code},${ref})`;
		delete cell.richText;
		putCell(sheet, row, col, cell);
	}
}

/**
 * Excel's Data > Subtotal: groups contiguous rows in `range` by `atEachChangeIn`, inserts group
 * summary rows and a Grand Total row with `=SUBTOTAL(code, ...)` formulas, and sets row outline
 * levels. One undo step.
 */
export function applySubtotal(
	ctx: EditContext,
	s: number,
	range: CellRange,
	options: SubtotalOptions,
): SubtotalResult {
	const sheet = sheetAt(ctx.workbook, s);
	const r = normalizeRange(range);
	const meta = resolveSubtotalFunction(options.fn);
	const { atEachChangeIn } = options;
	if (
		!Number.isInteger(atEachChangeIn) ||
		atEachChangeIn < r.start.col ||
		atEachChangeIn > r.end.col
	)
		throw new RangeError('Grouping column must be inside the subtotal range.');
	const valueCols = [...new Set(options.addSubtotalTo)];
	if (
		!valueCols.length ||
		valueCols.some((c) => !Number.isInteger(c) || c < r.start.col || c > r.end.col)
	)
		throw new RangeError('Subtotal columns must be non-empty and inside the subtotal range.');
	if (sheet.merges.some((m) => rangesIntersect(m, r)))
		throw new Error('Cannot apply subtotals to a range that contains merged cells.');

	const hasHeader = options.hasHeader ?? true;
	const replaceExisting = options.replaceExisting ?? true;
	const summaryBelow = options.summaryBelowData ?? true;
	const pageBreak = options.pageBreakBetweenGroups ?? false;
	const firstData = hasHeader ? r.start.row + 1 : r.start.row;
	if (firstData > r.end.row) return { range: r, subtotalRows: [], grandTotalRow: undefined };

	return ctx.run(
		'Subtotal',
		'structure',
		[{ kind: 'refs' }, { kind: 'sheet', sheet: s }],
		() => {
			let endRow = r.end.row;
			if (replaceExisting) endRow -= stripSubtotalsInPlace(ctx, sheet, r, firstData);
			if (firstData > endRow)
				return {
					range: { start: r.start, end: { row: endRow, col: r.end.col } },
					subtotalRows: [],
					grandTotalRow: undefined,
				};

			const groups = collectGroups(ctx, sheet, firstData, endRow, r, atEachChangeIn);
			if (!groups.length)
				return {
					range: { start: r.start, end: { row: endRow, col: r.end.col } },
					subtotalRows: [],
					grandTotalRow: undefined,
				};

			const { placed, grandTotalRow, finalEndRow } = insertSummaryRows(
				ctx,
				sheet,
				groups,
				firstData,
				endRow,
				summaryBelow,
			);
			writeSummaryRow(
				sheet,
				grandTotalRow,
				atEachChangeIn,
				meta.grandLabel,
				meta.code,
				valueCols,
				summaryBelow ? firstData : grandTotalRow + 1,
				summaryBelow ? grandTotalRow - 1 : finalEndRow,
			);
			setRowOutline(sheet, grandTotalRow, 0);

			for (let i = 0; i < placed.length; i++) {
				const g = placed[i]!;
				const label = g.label ? `${g.label} ${meta.suffix}` : `Blank ${meta.suffix}`;
				writeSummaryRow(
					sheet,
					g.subRow,
					atEachChangeIn,
					label,
					meta.code,
					valueCols,
					g.dataStart,
					g.dataEnd,
				);
				let maxDataLevel = 0;
				for (let row = g.dataStart; row <= g.dataEnd; row++) {
					const prev = sheet.rowInfo.get(row)?.outlineLevel ?? 0;
					const nextLevel = Math.min(MAX_OUTLINE_LEVEL, (prev || 1) + 1);
					setRowOutline(sheet, row, nextLevel);
					if (nextLevel > maxDataLevel) maxDataLevel = nextLevel;
				}
				setRowOutline(sheet, g.subRow, Math.max(1, maxDataLevel - 1));
				if (pageBreak) {
					const breakBefore = summaryBelow ? g.subRow + 1 : g.subRow;
					const shouldBreak = summaryBelow ? i < placed.length - 1 : i > 0;
					if (shouldBreak && breakBefore > 0) setPageBreak(ctx, s, 'row', breakBefore, true);
				}
			}

			return {
				range: { start: r.start, end: { row: finalEndRow, col: r.end.col } },
				subtotalRows: placed.map((g) => g.subRow),
				grandTotalRow,
			};
		},
		{ sheet: s, structural: true },
	);
}

/**
 * Excel's Data > Subtotal > Remove All: deletes all subtotal rows in `range` and clears row
 * outline levels across the range. One undo step.
 */
export function removeSubtotals(
	ctx: EditContext,
	s: number,
	range: CellRange,
	hasHeader = true,
): RemoveSubtotalsResult {
	const sheet = sheetAt(ctx.workbook, s);
	const r = normalizeRange(range);
	const firstData = hasHeader ? r.start.row + 1 : r.start.row;
	let hasAny = false;
	for (let row = r.start.row; row <= r.end.row; row++) {
		const info = sheet.rowInfo.get(row);
		if (
			(row >= firstData && isSubtotalRow(sheet, row, r.start.col, r.end.col)) ||
			info?.outlineLevel ||
			info?.collapsed
		) {
			hasAny = true;
			break;
		}
	}
	if (!hasAny) return { removed: 0, range: r };

	return ctx.run(
		'Remove subtotals',
		'structure',
		[{ kind: 'refs' }, { kind: 'sheet', sheet: s }],
		() => {
			const removed = stripSubtotalsInPlace(ctx, sheet, r, firstData);
			const endRow = Math.max(r.start.row, r.end.row - removed);
			return { removed, range: { start: r.start, end: { row: endRow, col: r.end.col } } };
		},
		{ sheet: s, structural: true },
	);
}
