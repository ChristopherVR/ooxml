// Rows, columns and merges of a shared sheet. Rows are keyed by row number and merge per row.
// Columns are keyed by span (`C:E`); concurrent splits of one span leave overlapping spans, which
// resolve column by column to the narrowest span covering it (the most specific edit). Merges are
// keyed by range; concurrent overlapping merges resolve to the first in row-major order.
import type * as Y from 'yjs';
import {
	type CellRange,
	columnIndex,
	columnLabel,
	formatRange,
	parseRange,
	rangesIntersect,
} from '../address';
import type { ColumnInfo, RowInfo, Worksheet } from '../model';
import { decodeColumn, decodeRow, encodeLine } from './codec';
import { setIfChanged } from './schema';

type KeyOf = (styleId: number | undefined) => string | undefined;
type IdOf = (key: string | undefined) => number | undefined;

export const rowKey = (row: number): string => String(row + 1);
export const spanKey = (min: number, max: number): string =>
	`${columnLabel(min)}:${columnLabel(max)}`;

export function parseRowKey(key: string): number | undefined {
	if (!/^\d{1,7}$/.test(key)) return undefined;
	const row = Number(key) - 1;
	return row >= 0 && row <= 1_048_575 ? row : undefined;
}

export function parseSpanKey(key: string): [number, number] | undefined {
	const match = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(key);
	if (!match) return undefined;
	const min = columnIndex(match[1] ?? '');
	const max = columnIndex(match[2] ?? '');
	return min >= 0 && max >= min && max <= 16_383 ? [min, max] : undefined;
}

/** Writes the sheet's rows, columns and merges where they differ. */
export function writeLines(
	sheet: Worksheet,
	parts: { rows: Y.Map<unknown>; cols: Y.Map<unknown>; merges: Y.Map<unknown> },
	keyOf: KeyOf,
): void {
	const rowKeys = new Set<string>();
	for (const [row, info] of sheet.rowInfo) {
		const key = rowKey(row);
		rowKeys.add(key);
		setIfChanged(parts.rows, key, encodeLine(info, keyOf));
	}
	for (const key of [...parts.rows.keys()]) if (!rowKeys.has(key)) parts.rows.delete(key);

	const colKeys = new Set<string>();
	for (const info of sheet.columns) {
		const key = spanKey(info.min, info.max);
		colKeys.add(key);
		setIfChanged(parts.cols, key, encodeLine(info, keyOf, true));
	}
	for (const key of [...parts.cols.keys()]) if (!colKeys.has(key)) parts.cols.delete(key);

	const mergeKeys = new Set(sheet.merges.map(formatRange));
	for (const key of mergeKeys) setIfChanged(parts.merges, key, true);
	for (const key of [...parts.merges.keys()]) if (!mergeKeys.has(key)) parts.merges.delete(key);
}

/** Applies changed row entries (all of them when `keys` is undefined). */
export function applyRows(
	sheet: Worksheet,
	rows: Y.Map<unknown>,
	idOf: IdOf,
	keys?: Iterable<string>,
): void {
	if (!keys) sheet.rowInfo.clear();
	for (const key of keys ?? rows.keys()) {
		const row = parseRowKey(key);
		if (row === undefined) continue;
		const info: RowInfo | undefined = rows.has(key) ? decodeRow(rows.get(key), idOf) : undefined;
		if (info) sheet.rowInfo.set(row, info);
		else sheet.rowInfo.delete(row);
	}
}

/** Rebuilds the column list from the shared spans. */
export function applyColumns(sheet: Worksheet, cols: Y.Map<unknown>, idOf: IdOf): void {
	const spans: ColumnInfo[] = [];
	for (const [key, raw] of cols.entries()) {
		const span = parseSpanKey(key);
		const info = span && decodeColumn(raw, span[0], span[1], idOf);
		if (info) spans.push(info);
	}
	const bounds = [...new Set(spans.flatMap((s) => [s.min, s.max + 1]))].sort((a, b) => a - b);
	const out: ColumnInfo[] = [];
	for (let i = 0; i + 1 < bounds.length; i++) {
		const min = bounds[i] ?? 0;
		const max = (bounds[i + 1] ?? 0) - 1;
		let pick: ColumnInfo | undefined;
		for (const span of spans) {
			if (span.min > min || span.max < max) continue;
			const width = span.max - span.min;
			const best = pick && pick.max - pick.min;
			if (!pick || best === undefined || width < best || (width === best && span.min < pick.min))
				pick = span;
		}
		if (!pick) continue;
		const last = out.at(-1);
		// Segments cut from the same span join again, so non-overlapping spans come back intact.
		if (last && last.max === min - 1 && spanOf.get(last) === pick) last.max = max;
		else {
			const piece = { ...pick, min, max };
			spanOf.set(piece, pick);
			out.push(piece);
		}
	}
	sheet.columns = out;
}
const spanOf = new WeakMap<ColumnInfo, ColumnInfo>();

/** Rebuilds the merges, dropping any that overlap an earlier one. */
export function applyMerges(sheet: Worksheet, merges: Y.Map<unknown>): void {
	const ranges = [...merges.keys()]
		.map(parseRange)
		.filter((r): r is CellRange => r !== undefined)
		.sort(
			(a, b) =>
				a.start.row - b.start.row ||
				a.start.col - b.start.col ||
				a.end.row - b.end.row ||
				a.end.col - b.end.col,
		);
	const kept: CellRange[] = [];
	for (const range of ranges)
		if (!kept.some((other) => rangesIntersect(other, range))) kept.push(range);
	sheet.merges = kept;
}
