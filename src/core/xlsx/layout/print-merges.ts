import type { CellRange } from '../address';

export interface PrintMergeSpan {
	range: CellRange;
	row: number;
	col: number;
	rowspan: number;
	colspan: number;
}
function lowerBound(values: readonly number[], target: number): number {
	let start = 0;
	let end = values.length;
	while (start < end) {
		const mid = (start + end) >>> 1;
		if (values[mid]! < target) start = mid + 1;
		else end = mid;
	}
	return start;
}

/** One merge visit and one indexed entry per covered visible cell; constant-time render lookups. */
export function createPrintMergeIndex(
	merges: readonly CellRange[],
	rows: readonly number[],
	columns: readonly number[],
) {
	const cells = new Map<number, Map<number, PrintMergeSpan>>();
	const stats = { mergesVisited: 0, cellsIndexed: 0, lookups: 0 };
	for (const range of merges) {
		stats.mergesVisited++;
		const rowStart = lowerBound(rows, range.start.row);
		const rowEnd = lowerBound(rows, range.end.row + 1);
		const colStart = lowerBound(columns, range.start.col);
		const colEnd = lowerBound(columns, range.end.col + 1);
		if (rowStart === rowEnd || colStart === colEnd) continue;
		const span = {
			range,
			row: rows[rowStart]!,
			col: columns[colStart]!,
			rowspan: rowEnd - rowStart,
			colspan: colEnd - colStart,
		};
		for (let r = rowStart; r < rowEnd; r++) {
			const row = rows[r]!;
			let line = cells.get(row);
			if (!line) {
				line = new Map();
				cells.set(row, line);
			}
			for (let c = colStart; c < colEnd; c++) {
				if (++stats.cellsIndexed > 2_000_000) throw new Error('Merged print ranges are too large.');
				if (!line.has(columns[c]!)) line.set(columns[c]!, span);
			}
		}
	}
	return {
		stats,
		at(row: number, col: number): PrintMergeSpan | undefined {
			stats.lookups++;
			return cells.get(row)?.get(col);
		},
	};
}
