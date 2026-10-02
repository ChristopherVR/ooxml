import { type CellRange, normalizeRange, rangesIntersect } from '../address.js';
import { forEachCellInRange, getCell } from '../cells.js';
import { type EditContext, pruneCell, sheetAt } from './context.js';
import { clearContents } from './cell-values.js';
import { patchRange } from './format.js';
import { cellRange } from './range-math.js';
import type { MergeMode } from './types.js';

/**
 * Merges cells. Only the top-left value survives (as in Excel); `center` also centres it and
 * `across` merges each row of the range separately. Merges overlapping the range are replaced.
 */
export function merge(ctx: EditContext, s: number, range: CellRange, mode: MergeMode): void {
	const sheet = sheetAt(ctx.workbook, s);
	const r = normalizeRange(range);
	if (r.start.row === r.end.row && r.start.col === r.end.col) return;
	if (sheet.tables.some((t) => rangesIntersect(t.range, r)))
		throw new Error('Cells inside a table cannot be merged.');
	const targets: CellRange[] =
		mode === 'across'
			? Array.from({ length: r.end.row - r.start.row + 1 }, (_v, i) =>
					cellRange(r.start.row + i, r.start.col, r.start.row + i, r.end.col),
				).filter((m) => m.start.col !== m.end.col)
			: [r];
	if (!targets.length) return;
	ctx.run(
		mode === 'center' ? 'Merge & Center' : mode === 'across' ? 'Merge Across' : 'Merge cells',
		'format',
		[{ kind: 'sheet', sheet: s }],
		() => {
			sheet.merges = sheet.merges.filter((m) => !rangesIntersect(m, r));
			for (const target of targets) {
				const anchor = target.start;
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
		[{ kind: 'sheet', sheet: s }],
		() => {
			sheet.merges = sheet.merges.filter((m) => !rangesIntersect(m, r));
		},
		{ sheet: s, ranges: [r] },
	);
}
