/** Largest zero-based row index in a SpreadsheetML sheet (1,048,576 rows). */
export const MAX_ROW = 1_048_575;
/** Largest zero-based column index in a SpreadsheetML sheet (16,384 columns, `XFD`). */
export const MAX_COL = 16_383;

/** A zero-based cell position. */
export interface CellAddress {
	row: number;
	col: number;
}

/** A rectangular, inclusive, zero-based range with `start` at the top left. */
export interface CellRange {
	start: CellAddress;
	end: CellAddress;
}

/** `0` -> `A`, `25` -> `Z`, `26` -> `AA`. */
export function columnLabel(col: number): string {
	let label = '';
	let n = col + 1;
	while (n > 0) {
		const rem = (n - 1) % 26;
		label = String.fromCharCode(65 + rem) + label;
		n = Math.floor((n - 1) / 26);
	}
	return label;
}

/** `A` -> `0`, `AA` -> `26`; case-insensitive. Returns `-1` for anything that is not letters. */
export function columnIndex(label: string): number {
	if (!/^[A-Za-z]{1,3}$/.test(label)) return -1;
	let n = 0;
	for (const ch of label.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
	return n - 1;
}

/** `{ row: 0, col: 0 }` -> `A1`. */
export const formatAddress = (address: CellAddress): string =>
	`${columnLabel(address.col)}${address.row + 1}`;

/** Parses `A1` or `$A$1`; returns `undefined` for malformed or out-of-grid references. */
export function parseAddress(ref: string): CellAddress | undefined {
	const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(ref.trim());
	if (!match) return undefined;
	const col = columnIndex(match[1] ?? '');
	const row = Number(match[2]) - 1;
	if (col < 0 || col > MAX_COL || row < 0 || row > MAX_ROW) return undefined;
	return { row, col };
}

/** Orders a range so `start` is the top-left corner. */
export function normalizeRange(range: CellRange): CellRange {
	return {
		start: {
			row: Math.min(range.start.row, range.end.row),
			col: Math.min(range.start.col, range.end.col),
		},
		end: {
			row: Math.max(range.start.row, range.end.row),
			col: Math.max(range.start.col, range.end.col),
		},
	};
}

/**
 * Parses `A1:B2`, a single cell `A1`, whole columns `A:C` or whole rows `1:3`. Whole-row and
 * whole-column ranges extend to the grid edge.
 */
export function parseRange(ref: string): CellRange | undefined {
	const parts = ref.trim().split(':');
	if (parts.length === 1) {
		const cell = parseAddress(parts[0] ?? '');
		return cell ? { start: cell, end: { ...cell } } : undefined;
	}
	if (parts.length !== 2) return undefined;
	const [a = '', b = ''] = parts.map((part) => part.replace(/\$/g, ''));
	if (/^[A-Za-z]{1,3}$/.test(a) && /^[A-Za-z]{1,3}$/.test(b)) {
		const c1 = columnIndex(a);
		const c2 = columnIndex(b);
		return normalizeRange({ start: { row: 0, col: c1 }, end: { row: MAX_ROW, col: c2 } });
	}
	if (/^\d+$/.test(a) && /^\d+$/.test(b)) {
		const r1 = Number(a) - 1;
		const r2 = Number(b) - 1;
		if (r1 < 0 || r2 < 0) return undefined;
		return normalizeRange({ start: { row: r1, col: 0 }, end: { row: r2, col: MAX_COL } });
	}
	const start = parseAddress(a);
	const end = parseAddress(b);
	return start && end ? normalizeRange({ start, end }) : undefined;
}

/** `A1:B2`, or `A1` when the range is one cell. */
export function formatRange(range: CellRange): string {
	const { start, end } = normalizeRange(range);
	const a = formatAddress(start);
	return start.row === end.row && start.col === end.col ? a : `${a}:${formatAddress(end)}`;
}

export const rangeContains = (range: CellRange, address: CellAddress): boolean =>
	address.row >= range.start.row &&
	address.row <= range.end.row &&
	address.col >= range.start.col &&
	address.col <= range.end.col;

export const rangesIntersect = (a: CellRange, b: CellRange): boolean =>
	a.start.row <= b.end.row &&
	b.start.row <= a.end.row &&
	a.start.col <= b.end.col &&
	b.start.col <= a.end.col;

/** Number of cells in a range. */
export const rangeSize = (range: CellRange): number =>
	(range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);

/** A numeric key for a cell, unique within a sheet; used for cell maps and dependency graphs. */
export const cellKey = (row: number, col: number): number => row * (MAX_COL + 1) + col;

/** The inverse of {@link cellKey}. */
export const keyToAddress = (key: number): CellAddress => ({
	row: Math.floor(key / (MAX_COL + 1)),
	col: key % (MAX_COL + 1),
});

/** Quotes a sheet name for use in a formula reference when it needs it (`'My Sheet'!A1`). */
export function quoteSheetName(name: string): string {
	return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !parseAddress(name)
		? name
		: `'${name.replace(/'/g, "''")}'`;
}
