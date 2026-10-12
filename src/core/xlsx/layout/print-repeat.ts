import type { CellRange } from '../address';
import type { Workbook } from '../model';
import type { GridMetrics } from './metrics';
import { printTitles } from './print-titles';

/** Leading whole-row/column titles, including titles outside the print area. */
export function printRepeat(
	workbook: Workbook,
	sheet: number,
	range: CellRange | undefined,
	metrics: GridMetrics,
): { rows: number[]; columns: number[]; width: number; height: number; warnings: string[] } {
	const titles = printTitles(workbook, sheet);
	const warnings: string[] = [];
	const rows: number[] = [];
	const columns: number[] = [];
	if (
		titles.unsupported ||
		(range &&
			((titles.rows && titles.rows.start > range.start.row) ||
				(titles.columns && titles.columns.start > range.start.col)))
	)
		warnings.push('Non-leading or unsupported print titles are preserved but not rendered.');
	else if (range) {
		if (
			(titles.rows ? titles.rows.end - titles.rows.start + 1 : 0) +
				(titles.columns ? titles.columns.end - titles.columns.start + 1 : 0) >
			10000
		)
			throw new Error('Print title ranges are too large for preview.');
		if (titles.rows)
			for (let row = titles.rows.start; row <= titles.rows.end; row++)
				if (!metrics.isRowHidden(row)) rows.push(row);
		if (titles.columns)
			for (let col = titles.columns.start; col <= titles.columns.end; col++)
				if (!metrics.isColHidden(col)) columns.push(col);
	}
	return {
		rows,
		columns,
		width: columns.reduce((sum, col) => sum + metrics.colWidth(col), 0),
		height: rows.reduce((sum, row) => sum + metrics.rowHeight(row), 0),
		warnings,
	};
}
