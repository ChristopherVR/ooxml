import { type CellRange, normalizeRange } from '../address.js';
import { forEachCellInRange } from '../cells.js';
import type { Workbook, Worksheet } from '../model.js';
import { type StylePatch, applyStylePatch } from '../styles.js';
import { editColumns } from './columns.js';
import {
	type EditContext,
	ensureCell,
	forEachPosition,
	isWholeColumns,
	isWholeRows,
	pruneCell,
	sheetAt,
} from './context.js';
import type { EditScope } from './history.js';

/** A memoised `applyStylePatch`, so a large range interns each distinct result once. */
export function stylePatcher(
	workbook: Workbook,
	patch: StylePatch,
): (id: number | undefined) => number {
	const memo = new Map<number, number>();
	return (id) => {
		const key = id ?? 0;
		let next = memo.get(key);
		if (next === undefined) {
			next = applyStylePatch(workbook, id, patch);
			memo.set(key, next);
		}
		return next;
	};
}

/** Applies a patch to every cell of a range in place (no undo step). */
export function patchRange(
	workbook: Workbook,
	sheet: Worksheet,
	range: CellRange,
	patch: StylePatch,
): void {
	patchRangeWith(sheet, range, stylePatcher(workbook, patch));
}

export function patchRangeWith(
	sheet: Worksheet,
	range: CellRange,
	patcher: (id: number | undefined) => number,
): void {
	const r = normalizeRange(range);
	const setId = (target: { styleId?: number }, id: number): void => {
		if (id) target.styleId = id;
		else delete target.styleId;
	};
	if (isWholeColumns(r) || isWholeRows(r)) {
		if (isWholeColumns(r)) {
			const cols: number[] = [];
			for (let c = r.start.col; c <= r.end.col; c++) cols.push(c);
			editColumns(sheet, cols, (info) => setId(info, patcher(info.styleId)));
		}
		if (isWholeRows(r) && !isWholeColumns(r))
			for (let row = r.start.row; row <= r.end.row; row++) {
				const info = sheet.rowInfo.get(row) ?? {};
				setId(info, patcher(info.styleId));
				sheet.rowInfo.set(row, info);
			}
		const stored: [number, number][] = [];
		forEachCellInRange(sheet, r, (cell, row, col) => {
			setId(cell, patcher(cell.styleId));
			stored.push([row, col]);
		});
		for (const [row, col] of stored) pruneCell(sheet, row, col);
		return;
	}
	forEachPosition(sheet, r, (row, col) => {
		const cell = ensureCell(sheet, row, col);
		setId(cell, patcher(cell.styleId));
		pruneCell(sheet, row, col);
	});
}

export function formatScope(index: number, ranges: CellRange[]): EditScope {
	const whole = ranges.some((r) => isWholeColumns(r) || isWholeRows(r));
	return whole ? { kind: 'sheet', sheet: index } : { kind: 'cells', sheet: index, ranges };
}

export function applyStyle(
	ctx: EditContext,
	s: number,
	ranges: CellRange[],
	patch: StylePatch,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const normalized = ranges.map(normalizeRange);
	if (!normalized.length) return;
	const patcher = stylePatcher(ctx.workbook, patch);
	ctx.run(
		'Format cells',
		'format',
		[formatScope(s, normalized)],
		() => {
			for (const range of normalized) patchRangeWith(sheet, range, patcher);
		},
		{ sheet: s, ranges: normalized },
	);
}
