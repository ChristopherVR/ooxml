import type { CellDifference, CellDifferenceRange } from './types';

const KIND_ORDER: Record<CellDifference['kind'], number> = {
	enteredValue: 0,
	formula: 1,
	calculatedValue: 2,
	format: 3,
};

interface Run {
	key: string;
	first: CellDifference;
	start: number;
	end: number;
}

/**
 * Groups row-major cell differences into rectangles of neighbouring cells with the same kind (and,
 * for formats, the same changed attributes): runs along each row, then runs with the same columns
 * on consecutive rows. Linear in the number of differences, apart from sorting the result.
 */
export function groupCellDifferences(cells: readonly CellDifference[]): CellDifferenceRange[] {
	const out: CellDifferenceRange[] = [];
	let above = new Map<string, CellDifferenceRange>();
	let current = new Map<string, CellDifferenceRange>();
	let runs = new Map<string, Run>();
	let row = -2;

	const close = (run: Run) => {
		const key = `${run.key}\n${run.start}:${run.end}`;
		const width = run.end - run.start + 1;
		const open = above.get(key);
		if (open) {
			open.range.end.row = row;
			open.count += width;
			current.set(key, open);
			return;
		}
		const range: CellDifferenceRange = {
			kind: run.first.kind,
			range: { start: { row, col: run.start }, end: { row, col: run.end } },
			count: width,
		};
		if (run.first.attributes) range.attributes = run.first.attributes;
		out.push(range);
		current.set(key, range);
	};
	const endRow = () => {
		for (const run of runs.values()) close(run);
		runs = new Map();
	};

	for (const cell of cells) {
		const { row: r, col } = cell.address;
		if (r !== row) {
			endRow();
			above = r === row + 1 ? current : new Map();
			current = new Map();
			row = r;
		}
		const key = cell.attributes ? `${cell.kind}:${cell.attributes.join()}` : cell.kind;
		const run = runs.get(key);
		if (run && run.end === col - 1) run.end = col;
		else {
			if (run) close(run);
			runs.set(key, { key, first: cell, start: col, end: col });
		}
	}
	endRow();
	return out.sort(
		(a, b) =>
			a.range.start.row - b.range.start.row ||
			a.range.start.col - b.range.start.col ||
			KIND_ORDER[a.kind] - KIND_ORDER[b.kind],
	);
}
