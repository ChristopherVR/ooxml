// The visible cells of a table as an HTML table draws them: row spans for vertical merges, and the
// resolved borders, shading and margins (direct formatting over the table style) in the attribute
// shape `tableStyle` and `tableCellStyle` read. The editor's table nodes and the read-only header,
// footer and note renderers share it, so a story table draws its borders like a body table.
import type { Table, TableCell, TableStyleCatalog } from '../index';
import { resolveCellVisuals } from '../index';
import type { ThemeCatalog } from '../theme-model';
import { continuationsOf, rowSpanOf, tableGrid } from '../table-grid';

/** One cell an HTML table shows (vertical merge continuations are covered by `rowspan`). */
export interface TableCellView {
	cell: TableCell;
	/** First grid column. */
	column: number;
	colspan: number;
	rowspan: number;
	/** The continuation cells below a vertical merge start, top to bottom. */
	continues: TableCell[];
	/** `tableCellStyle` input: resolved visuals, direct margins and the preferred width. */
	attrs: {
		widthTwips: number | null;
		verticalAlign: string | null;
		shadingFill: string | null;
		borders: string | null;
		margins: string | null;
	};
}

/** `tableStyle` input for a table: its width, alignment, indent, borders and cell margins. */
export function tableViewAttrs(block: Table): {
	widthTwips: number | null;
	widthPercent: number | null;
	alignment: string | null;
	indentTwips: number | null;
	borders: string | null;
	cellMargins: string | null;
} {
	return {
		widthTwips: block.widthTwips ?? null,
		widthPercent: block.widthPercent ?? null,
		alignment: block.alignment ?? null,
		indentTwips: block.indentTwips ?? null,
		borders: block.borders ? JSON.stringify(block.borders) : null,
		cellMargins: block.cellMargins ? JSON.stringify(block.cellMargins) : null,
	};
}

/** The cells of each row an HTML table shows, with their resolved visuals. */
export function tableCellViews(
	block: Table,
	tableStyles?: TableStyleCatalog,
	theme?: ThemeCatalog,
): TableCellView[][] {
	const columnCount = Math.max(
		1,
		...block.rows.map((row) => row.reduce((sum, cell) => sum + (cell.gridSpan ?? 1), 0)),
	);
	const layout = tableGrid(block);
	return block.rows.map((row, rowIndex) => {
		const views: TableCellView[] = [];
		let gridColumn = 0;
		row.forEach((cell, columnIndex) => {
			const column = gridColumn;
			const colspan = cell.gridSpan ?? 1;
			gridColumn += colspan;
			if (cell.verticalMerge === 'continue') return;
			let rowspan = 1;
			const entry = layout?.rows[rowIndex]?.[columnIndex];
			const continues = layout && entry ? continuationsOf(layout, entry) : [];
			if (layout && entry) rowspan = rowSpanOf(layout, entry);
			else if (cell.verticalMerge === 'restart') {
				let next = rowIndex + 1;
				while (block.rows[next]?.[columnIndex]?.verticalMerge === 'continue') {
					rowspan++;
					next++;
				}
			}
			const visuals = resolveCellVisuals(
				block,
				cell,
				{
					row: rowIndex,
					lastRow: rowIndex + rowspan - 1,
					column,
					lastColumn: column + colspan - 1,
					rowCount: block.rows.length,
					columnCount,
				},
				tableStyles,
				theme,
			);
			views.push({
				cell,
				column,
				colspan,
				rowspan,
				continues: continues.map((continued) => continued.cell),
				attrs: {
					widthTwips: cell.widthTwips ?? null,
					verticalAlign: cell.verticalAlign ?? null,
					shadingFill: visuals.shadingFill ?? null,
					borders: visuals.borders ? JSON.stringify(visuals.borders) : null,
					margins: cell.margins ? JSON.stringify(cell.margins) : null,
				},
			});
		});
		return views;
	});
}
