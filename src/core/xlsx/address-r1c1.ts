import { MAX_COL, MAX_ROW, normalizeRange, type CellAddress, type CellRange } from './address';

interface Corner {
	kind: 'cell' | 'row' | 'col';
	range: CellRange;
}

const dimension = (text: string | undefined, base: number, size: number): number | undefined => {
	if (!text) return base;
	if (text.startsWith('[')) {
		const offset = Number(text.slice(1, -1));
		// Excel wraps relative coordinates at sheet edges, but rejects a full-grid offset.
		if (Math.abs(offset) >= size) return undefined;
		return (((base + offset) % size) + size) % size;
	}
	const at = Number(text) - 1;
	return at >= 0 && at < size ? at : undefined;
};

function corner(text: string, base: CellAddress): Corner | undefined {
	const cell = /^R(\[[+-]?\d+\]|\d+)?C(\[[+-]?\d+\]|\d+)?$/i.exec(text);
	if (cell) {
		const row = dimension(cell[1], base.row, MAX_ROW + 1);
		const col = dimension(cell[2], base.col, MAX_COL + 1);
		if (row === undefined || col === undefined) return undefined;
		return { kind: 'cell', range: { start: { row, col }, end: { row, col } } };
	}
	const rowMatch = /^R(\[[+-]?\d+\]|\d+)?$/i.exec(text);
	if (rowMatch) {
		const row = dimension(rowMatch[1], base.row, MAX_ROW + 1);
		if (row === undefined) return undefined;
		return { kind: 'row', range: { start: { row, col: 0 }, end: { row, col: MAX_COL } } };
	}
	const colMatch = /^C(\[[+-]?\d+\]|\d+)?$/i.exec(text);
	if (colMatch) {
		const col = dimension(colMatch[1], base.col, MAX_COL + 1);
		if (col === undefined) return undefined;
		return { kind: 'col', range: { start: { row: 0, col }, end: { row: MAX_ROW, col } } };
	}
	return undefined;
}

/**
 * Parses an unqualified R1C1 cell/range or whole rows/columns relative to a zero-based base.
 * Absolute indexes are one-based. Relative indexes wrap at sheet edges, as in Excel INDIRECT.
 * Malformed references, mixed endpoint kinds and out-of-grid absolute indexes return undefined.
 */
export function parseR1C1Range(text: string, base: CellAddress): CellRange | undefined {
	if (
		!Number.isInteger(base.row) ||
		!Number.isInteger(base.col) ||
		base.row < 0 ||
		base.row > MAX_ROW ||
		base.col < 0 ||
		base.col > MAX_COL
	)
		return undefined;
	const parts = text.trim().split(':');
	if (parts.length > 2) return undefined;
	const a = corner(parts[0] ?? '', base);
	const b = parts.length === 1 ? a : corner(parts[1] ?? '', base);
	if (!a || !b || a.kind !== b.kind) return undefined;
	return normalizeRange({ start: a.range.start, end: b.range.end });
}
