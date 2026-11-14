/**
 * table-cells-format.ts: the Home tab's character formatting for a selected
 * table, decided once for every binding. With a cell range (or the table
 * frame) selected, Bold, Italic, Underline, Strikethrough, a size, a font or
 * a colour applies to every run of every selected cell; the toggles read
 * their state from the first selected cell, as PowerPoint's ribbon shows it.
 *
 * Sizes are points here (cell units), the size the ribbon's box shows.
 *
 * @module render/table-cells-format
 */
import type { PptxTableCell, TablePptxElement } from 'ooxml-core/pptx';
import {
	applyTableCellsTextStyle,
	type TableCellRef,
	type TableCellTextStyleUpdate,
} from 'ooxml-core/pptx/editor/render/table-cells-text-style';

/** The character format the ribbon shows for selected table cells. */
export interface TableCellsTextFormat {
	bold: boolean;
	italic: boolean;
	underline: boolean;
	strikethrough: boolean;
	/** Points, or `undefined` when neither the run nor the cell sets one. */
	fontSize: number | undefined;
	fontFamily: string | undefined;
	color: string | undefined;
}

function firstTarget(
	element: TablePptxElement,
	cells: readonly TableCellRef[] | undefined,
): PptxTableCell | undefined {
	const rows = element.tableData?.rows ?? [];
	if (cells && cells.length > 0) {
		const ordered = [...cells].sort((a, b) => a.row - b.row || a.col - b.col);
		return rows[ordered[0]!.row]?.cells[ordered[0]!.col];
	}
	for (const row of rows) {
		const cell = row.cells.find((item) => !item.hMerge && !item.vMerge);
		if (cell) {
			return cell;
		}
	}
	return undefined;
}

/** The format of the first selected cell (its first run, else the cell style). */
export function tableCellsTextFormat(
	element: TablePptxElement,
	cells?: readonly TableCellRef[],
): TableCellsTextFormat {
	const cell = firstTarget(element, cells);
	const run = cell?.textRuns?.find((item) => !item.isParagraphBreak && !item.isLineBreak);
	const style = cell?.style;
	return {
		bold: Boolean(run?.bold ?? style?.bold),
		italic: Boolean(run?.italic ?? style?.italic),
		underline: Boolean(run?.underline ?? style?.underline),
		strikethrough: Boolean(run?.strikethrough),
		fontSize: run?.fontSize ?? style?.fontSize ?? element.tableData?.defaultCellFontSize,
		fontFamily: run?.fontFamily ?? style?.fontFamily,
		color: run?.color ?? style?.color,
	};
}

/**
 * The element patch that applies `update` to the selected cells (every cell
 * when `cells` is empty or omitted), or an empty patch when nothing changes.
 */
export function tableCellsTextStylePatch(
	element: TablePptxElement,
	cells: readonly TableCellRef[] | undefined,
	update: TableCellTextStyleUpdate,
): Partial<TablePptxElement> {
	const next = applyTableCellsTextStyle(element, cells, update);
	if (next === element) {
		return {};
	}
	return { tableData: next.tableData, ...(next.rawXml ? { rawXml: next.rawXml } : {}) };
}
