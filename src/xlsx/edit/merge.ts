import { type CellRange, normalizeRange, rangesIntersect } from '../address.js';
import { forEachCellInRange, getCell, putCell } from '../cells.js';
import type { Cell, Worksheet } from '../model.js';
import { type EditContext, pruneCell, sheetAt } from './context.js';
import { clearContents } from './cell-values.js';
import { isSpilledCell } from './deps.js';
import { patchRange } from './format.js';
import { cellRange } from './range-math.js';
import type { MergeMode } from './types.js';

const hasContent = (cell: Cell | undefined): cell is Cell =>
	!!cell &&
	!isSpilledCell(cell) &&
	(cell.formula !== undefined || (cell.value !== null && cell.value !== ''));

function mergeTargets(r: CellRange, mode: MergeMode): CellRange[] {
	return mode === 'across'
		? Array.from({ length: r.end.row - r.start.row + 1 }, (_v, i) =>
				cellRange(r.start.row + i, r.start.col, r.start.row + i, r.end.col),
			).filter((m) => m.start.col !== m.end.col)
		: [r];
}

/** The cells of a merge target that hold a value or formula, in row-major order. */
function filledCells(sheet: Worksheet, target: CellRange): [number, number][] {
	const filled: [number, number][] = [];
	forEachCellInRange(sheet, target, (cell, row, col) => {
		if (hasContent(cell)) filled.push([row, col]);
	});
	return filled.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
}

/**
 * Whether merging would discard data: some merge target holds more than one non-empty cell.
 * Excel warns "Merging cells only keeps the upper-left value and discards other values" then.
 */
export function mergeWouldDiscard(
	sheet: Worksheet,
	range: CellRange,
	mode: MergeMode = 'merge',
): boolean {
	return mergeTargets(normalizeRange(range), mode).some((t) => filledCells(sheet, t).length > 1);
}

/**
 * Merges cells. The first non-empty cell (row by row) survives in the top-left corner and the
 * others are cleared, as in Excel 16 (which moves that cell's value, formula and format to the
 * corner when the corner is empty); `center` also centres it and `across` merges each row of
 * the range separately. Merges overlapping the range are replaced.
 */
export function merge(ctx: EditContext, s: number, range: CellRange, mode: MergeMode): void {
	const sheet = sheetAt(ctx.workbook, s);
	const r = normalizeRange(range);
	if (r.start.row === r.end.row && r.start.col === r.end.col) return;
	if (sheet.tables.some((t) => rangesIntersect(t.range, r)))
		throw new Error('Cells inside a table cannot be merged.');
	const targets = mergeTargets(r, mode);
	if (!targets.length) return;
	ctx.run(
		mode === 'center' ? 'Merge & Center' : mode === 'across' ? 'Merge Across' : 'Merge cells',
		'format',
		[
			{ kind: 'cells', sheet: s, ranges: [r] },
			{ kind: 'parts', sheet: s, parts: ['merges', 'columns', 'rowInfo'] },
		],
		() => {
			sheet.merges = sheet.merges.filter((m) => !rangesIntersect(m, r));
			for (const target of targets) {
				const anchor = target.start;
				const [first] = filledCells(sheet, target);
				if (first && (first[0] !== anchor.row || first[1] !== anchor.col)) {
					const source = getCell(sheet, first[0], first[1]);
					const corner = getCell(sheet, anchor.row, anchor.col);
					if (source) {
						const moved: Cell = structuredClone(source);
						if (moved.styleId === undefined && corner?.styleId !== undefined)
							moved.styleId = corner.styleId;
						delete moved.arrayRange;
						putCell(sheet, anchor.row, anchor.col, moved);
					}
				}
				const doomed: [number, number][] = [];
				forEachCellInRange(sheet, target, (_cell, row, col) => {
					if (row !== anchor.row || col !== anchor.col) doomed.push([row, col]);
				});
				for (const [row, col] of doomed) {
					const cell = getCell(sheet, row, col);
					if (cell) clearContents(cell);
					pruneCell(sheet, row, col);
				}
				sheet.merges.push(target);
			}
			if (mode === 'center')
				patchRange(ctx.workbook, sheet, r, { alignment: { horizontal: 'center' } });
		},
		{ sheet: s, ranges: [r] },
	);
}

/** Removes every merge that overlaps `range`. */
export function unmerge(ctx: EditContext, s: number, range: CellRange): void {
	const sheet = sheetAt(ctx.workbook, s);
	const r = normalizeRange(range);
	if (!sheet.merges.some((m) => rangesIntersect(m, r))) return;
	ctx.run(
		'Unmerge cells',
		'format',
		[{ kind: 'parts', sheet: s, parts: ['merges'] }],
		() => {
			sheet.merges = sheet.merges.filter((m) => !rangesIntersect(m, r));
		},
		{ sheet: s, ranges: [r] },
	);
}
