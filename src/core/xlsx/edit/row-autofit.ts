import { type CellRange, normalizeRange } from '../address.js';
import type { MeasureText } from '../layout/autofit.js';
import type { Worksheet } from '../model.js';
import { approximateMeasure, autoFitRowHeight } from '../layout/row-autofit.js';
import { type EditContext, sheetAt } from './context.js';

/** Excel's largest row height in points. */
const MAX_HEIGHT = 409;

/**
 * Re-fits rows that have no custom height after an edit changed their text, wrapping or font
 * size, as Excel grows (and shrinks back) such rows automatically. Rows with a custom height or
 * hidden rows are left alone. Nothing happens when no height changes; otherwise the change is
 * recorded as its own undo entry, which joins the surrounding step when called inside an edit or
 * a batch (call it from inside `ctx.run` so the edit and the row heights undo together). Returns
 * the rows whose height changed. Without `measure`, text width is approximated from the font
 * size, so wrapped line counts are estimates.
 */
export function autoGrowRows(
	ctx: EditContext,
	s: number,
	rows: Iterable<number>,
	measure: MeasureText = approximateMeasure,
): number[] {
	const sheet = sheetAt(ctx.workbook, s);
	const updates: [number, number | undefined][] = [];
	for (const row of new Set(rows)) {
		const info = sheet.rowInfo.get(row);
		if (info?.customHeight || info?.hidden) continue;
		const fit = Math.min(MAX_HEIGHT, autoFitRowHeight(ctx.workbook, s, row, measure));
		const next = Math.abs(fit - sheet.defaultRowHeight) < 1e-9 ? undefined : fit;
		const current = info?.height;
		if (current === next) continue;
		if (current !== undefined && next !== undefined && Math.abs(current - next) < 1e-9) continue;
		updates.push([row, next]);
	}
	if (!updates.length) return [];
	ctx.run(
		'Row height',
		'view',
		[{ kind: 'sheet', sheet: s }],
		() => {
			for (const [row, height] of updates) {
				const info = { ...sheet.rowInfo.get(row) };
				if (height === undefined) delete info.height;
				else info.height = height;
				if (Object.keys(info).length) sheet.rowInfo.set(row, info);
				else sheet.rowInfo.delete(row);
			}
		},
		{ sheet: s, structural: true },
	);
	return updates.map(([row]) => row);
}

/** Largest number of rows one edit re-fits. */
const MAX_REFIT_ROWS = 10_000;

/** The rows of `ranges` that hold cells or a non-custom height (candidates for re-fitting). */
export function rowsToRefit(sheet: Worksheet, ranges: readonly CellRange[]): number[] {
	const rows = new Set<number>();
	const wanted = (row: number): boolean =>
		sheet.rows.has(row) || sheet.rowInfo.get(row)?.height !== undefined;
	for (const range of ranges) {
		const r = normalizeRange(range);
		if (r.end.row - r.start.row < MAX_REFIT_ROWS) {
			for (let row = r.start.row; row <= r.end.row; row++) if (wanted(row)) rows.add(row);
		} else {
			for (const row of new Set([...sheet.rows.keys(), ...sheet.rowInfo.keys()]))
				if (row >= r.start.row && row <= r.end.row && wanted(row)) rows.add(row);
		}
		if (rows.size >= MAX_REFIT_ROWS) break;
	}
	return [...rows].slice(0, MAX_REFIT_ROWS).sort((a, b) => a - b);
}
