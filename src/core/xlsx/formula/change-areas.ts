import type { CellPosition } from './engine-types';
import type { Area } from './values';

/**
 * Changed cells as rectangles, for dependency queries that take a range: runs along a row, then
 * runs over the same columns on consecutive rows. Cells in row-major order (as a bulk edit writes
 * them) give a handful of areas; any order still gives areas whose union is exactly the cells.
 */
export function changeAreas(changes: readonly CellPosition[]): Area[] {
	const out: Area[] = [];
	// The rectangle each column span last grew, by sheet and span.
	const open = new Map<string, Area>();
	let run: Area | undefined;
	const close = (area: Area): void => {
		const { start, end } = area.range;
		const key = `${area.sheet}:${start.col}:${end.col}`;
		const above = open.get(key);
		if (above && above.range.end.row === start.row - 1) {
			above.range.end.row = end.row;
			return;
		}
		out.push(area);
		open.set(key, area);
	};
	for (const { sheet, row, col } of changes) {
		if (run && run.sheet === sheet && run.range.start.row === row) {
			if (col === run.range.end.col + 1) {
				run.range.end.col = col;
				continue;
			}
			if (col >= run.range.start.col && col <= run.range.end.col) continue;
		}
		if (run) close(run);
		run = { sheet, range: { start: { row, col }, end: { row, col } } };
	}
	if (run) close(run);
	return out;
}
