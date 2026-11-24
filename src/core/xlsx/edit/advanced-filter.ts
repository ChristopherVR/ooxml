import { type CellAddress, type CellRange, MAX_COL, MAX_ROW, normalizeRange } from '../address';
import { deleteCell, getCell, putCell } from '../cells';
import type { RowInfo, Workbook, Worksheet } from '../model';
import { compileAdvancedCriteria } from './advanced-filter-criteria';
import { cloneCell } from './clone-cell';
import { type EditContext, displayText, sheetAt } from './context';

export type AdvancedFilterAction = 'filter-in-place' | 'copy-to-location';

export interface AdvancedFilterOptions {
	/** List range including its header row (first row). */
	listRange: CellRange;
	/** Optional criteria range (header row + 1 or more criteria rows; AND across columns, OR across rows). */
	criteriaRange?: CellRange;
	/** Optional zero-based sheet index for `criteriaRange` (defaults to the target sheet). */
	criteriaSheet?: number;
	/** Filter the list in place (hiding non-matching rows) or copy matching rows to `copyTo`. */
	action: AdvancedFilterAction;
	/** Destination top-left cell or header row range when `action === 'copy-to-location'`. */
	copyTo?: CellAddress | CellRange;
	/** Optional zero-based sheet index for `copyTo` (defaults to the target sheet). */
	copyToSheet?: number;
	/** Keep only the first occurrence of duplicate matching rows (default false). */
	uniqueOnly?: boolean;
}

export interface AdvancedFilterResult {
	/** Number of matching data rows (after deduplication when `uniqueOnly` is true). */
	matched: number;
	/** Total number of data rows evaluated in `listRange`. */
	total: number;
	/** Destination range written for `'copy-to-location'`, or `listRange` for `'filter-in-place'`. */
	range: CellRange;
}

function validateAddress(addr: CellAddress, label: string): void {
	if (
		!Number.isInteger(addr.row) ||
		!Number.isInteger(addr.col) ||
		addr.row < 0 ||
		addr.row > MAX_ROW ||
		addr.col < 0 ||
		addr.col > MAX_COL
	)
		throw new RangeError(`${label} must be a valid cell address.`);
}

function setRowVisibility(sheet: Worksheet, row: number, hidden: boolean): void {
	const existing = sheet.rowInfo.get(row);
	const info: RowInfo = existing ? { ...existing } : {};
	if (hidden) {
		info.hidden = true;
		info.filteredOut = true;
	} else {
		delete info.hidden;
		delete info.filteredOut;
		delete info.manuallyHidden;
	}
	if (Object.keys(info).length) sheet.rowInfo.set(row, info);
	else sheet.rowInfo.delete(row);
}

function resolveCopyColumns(
	workbook: Workbook,
	listSheet: Worksheet,
	listRange: CellRange,
	destSheet: Worksheet,
	copyTo: CellAddress | CellRange,
): { topLeft: CellAddress; cols: number[]; preserveHeaders: boolean } {
	const isRange = 'start' in copyTo;
	const norm = isRange ? normalizeRange(copyTo) : { start: copyTo, end: copyTo };
	validateAddress(norm.start, 'Advanced Filter copyTo');
	const topLeft = norm.start;

	const headerMap = new Map<string, number>();
	const allCols: number[] = [];
	for (let c = listRange.start.col; c <= listRange.end.col; c++) {
		allCols.push(c);
		const key = displayText(workbook, getCell(listSheet, listRange.start.row, c))
			.trim()
			.toLowerCase();
		if (key && !headerMap.has(key)) headerMap.set(key, c);
	}

	// Detect pre-populated header subset at `copyTo`.
	const requested: number[] = [];
	if (isRange && norm.end.col > norm.start.col) {
		let anyNonEmpty = false;
		for (let c = norm.start.col; c <= norm.end.col; c++) {
			const text = displayText(workbook, getCell(destSheet, topLeft.row, c)).trim();
			if (!text) continue;
			anyNonEmpty = true;
			const srcCol = headerMap.get(text.toLowerCase());
			if (srcCol === undefined)
				throw new Error(`Advanced Filter copyTo header "${text}" was not found in the list range.`);
			requested.push(srcCol);
		}
		if (anyNonEmpty) return { topLeft, cols: requested, preserveHeaders: true };
	} else {
		const firstText = displayText(workbook, getCell(destSheet, topLeft.row, topLeft.col)).trim();
		if (firstText && headerMap.has(firstText.toLowerCase())) {
			for (let c = topLeft.col; c <= MAX_COL; c++) {
				const text = displayText(workbook, getCell(destSheet, topLeft.row, c)).trim();
				if (!text) break;
				const srcCol = headerMap.get(text.toLowerCase());
				if (srcCol === undefined)
					throw new Error(
						`Advanced Filter copyTo header "${text}" was not found in the list range.`,
					);
				requested.push(srcCol);
			}
			return { topLeft, cols: requested, preserveHeaders: true };
		}
	}

	return { topLeft, cols: allCols, preserveHeaders: false };
}

function rowKey(
	workbook: Workbook,
	sheet: Worksheet,
	row: number,
	cols: readonly number[],
): string {
	return cols
		.map((col) => displayText(workbook, getCell(sheet, row, col)).toLowerCase())
		.join('\u0000');
}

function copyRowValues(
	srcSheet: Worksheet,
	srcRow: number,
	srcCols: readonly number[],
	destSheet: Worksheet,
	destRow: number,
	destColStart: number,
): void {
	for (let i = 0; i < srcCols.length; i++) {
		const srcCell = getCell(srcSheet, srcRow, srcCols[i]!);
		const destCol = destColStart + i;
		if (!srcCell || srcCell.value === null) {
			deleteCell(destSheet, destRow, destCol);
			continue;
		}
		const copy = cloneCell(srcCell);
		delete copy.formula;
		delete copy.arrayRange;
		delete (copy as { spillAnchor?: unknown }).spillAnchor;
		putCell(destSheet, destRow, destCol, copy);
	}
}

/**
 * Excel's Data > Advanced Filter: evaluates an optional criteria range (columns ANDed within each
 * criteria row, criteria rows ORed) over `listRange` (header row + data rows). Either hides
 * non-matching rows in place or copies headers (or a requested header subset at `copyTo`) and
 * matching rows to `copyTo`, with optional deduplication (`uniqueOnly`). One undo step.
 */
export function applyAdvancedFilter(
	ctx: EditContext,
	s: number,
	options: AdvancedFilterOptions,
): AdvancedFilterResult {
	const listSheet = sheetAt(ctx.workbook, s);
	const listRange = normalizeRange(options.listRange);
	const criteriaSheetIdx = options.criteriaSheet ?? s;
	const criteriaSheet = sheetAt(ctx.workbook, criteriaSheetIdx);
	const matches = options.criteriaRange
		? compileAdvancedCriteria(
				ctx.workbook,
				listSheet,
				listRange,
				criteriaSheet,
				normalizeRange(options.criteriaRange),
			)
		: () => true;

	const firstData = listRange.start.row + 1;
	const total = Math.max(0, listRange.end.row - firstData + 1);
	const uniqueOnly = options.uniqueOnly ?? false;

	if (options.action === 'filter-in-place') {
		const allCols: number[] = [];
		for (let c = listRange.start.col; c <= listRange.end.col; c++) allCols.push(c);
		const keepSet = new Set<number>();
		const seen = new Set<string>();
		for (let row = firstData; row <= listRange.end.row; row++) {
			if (!matches(row)) continue;
			if (uniqueOnly) {
				const key = rowKey(ctx.workbook, listSheet, row, allCols);
				if (seen.has(key)) continue;
				seen.add(key);
			}
			keepSet.add(row);
		}
		ctx.run(
			'Advanced Filter',
			'view',
			[{ kind: 'sheet', sheet: s }],
			() => {
				for (let row = firstData; row <= listRange.end.row; row++)
					setRowVisibility(listSheet, row, !keepSet.has(row));
			},
			{ sheet: s, structural: true },
		);
		return { matched: keepSet.size, total, range: listRange };
	}

	if (!options.copyTo)
		throw new RangeError('Advanced Filter with action "copy-to-location" requires copyTo.');
	const destSheetIdx = options.copyToSheet ?? s;
	const destSheet = sheetAt(ctx.workbook, destSheetIdx);
	const { topLeft, cols, preserveHeaders } = resolveCopyColumns(
		ctx.workbook,
		listSheet,
		listRange,
		destSheet,
		options.copyTo,
	);

	const matchedRows: number[] = [];
	const seen = new Set<string>();
	for (let row = firstData; row <= listRange.end.row; row++) {
		if (!matches(row)) continue;
		if (uniqueOnly) {
			const key = rowKey(ctx.workbook, listSheet, row, cols);
			if (seen.has(key)) continue;
			seen.add(key);
		}
		matchedRows.push(row);
	}

	const destRange: CellRange = {
		start: { row: topLeft.row, col: topLeft.col },
		end: { row: topLeft.row + matchedRows.length, col: topLeft.col + cols.length - 1 },
	};
	if (destRange.end.row > MAX_ROW || destRange.end.col > MAX_COL)
		throw new RangeError('Advanced Filter output exceeds the worksheet grid.');

	ctx.run(
		'Advanced Filter',
		'cells',
		[{ kind: 'cells', sheet: destSheetIdx, ranges: [destRange] }],
		() => {
			if (!preserveHeaders)
				copyRowValues(listSheet, listRange.start.row, cols, destSheet, topLeft.row, topLeft.col);
			for (let i = 0; i < matchedRows.length; i++)
				copyRowValues(
					listSheet,
					matchedRows[i]!,
					cols,
					destSheet,
					topLeft.row + 1 + i,
					topLeft.col,
				);
		},
		{ sheet: destSheetIdx, ranges: [destRange] },
	);

	return { matched: matchedRows.length, total, range: destRange };
}
