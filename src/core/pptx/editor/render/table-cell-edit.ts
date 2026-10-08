/**
 * table-cell-edit.ts: Pure, immutable table-cell editing helpers shared by
 * every binding so inline cell editing is not reimplemented per framework.
 *
 * All functions return new objects and leave the input element unchanged.
 *
 * @module render/table-cell-edit
 */
import type { PptxTableCell, PptxTableCellStyle, TablePptxElement } from 'ooxml-core/pptx';

/**
 * Replace one cell's plain text, DROPPING the per-run model that described the
 * text it used to hold.
 *
 * `PptxTableCell` carries the same content twice: `text`, a flat `\n`-joined
 * string, and `textRuns`, the styled run sequence core parses out of the cell's
 * `a:txBody`. Every binding's table renderer paints `textRuns` when it is
 * present and only falls back to `text` when it is not, so spreading a cell and
 * overwriting `text` alone leaves the OLD wording painted: the edit lands in the
 * model, `updateCellTextInRawXml` lands it in the markup, and the canvas keeps
 * showing what was there before. That is precisely what `desktop-manipulation`
 * and `mobile-table` caught in four of the five bindings; Vanilla passed only
 * because its editor had spelled out `textRuns: undefined` locally, which is the
 * usual signal that a shared decision was missing.
 *
 * Dropping the runs rather than rebuilding them is also what the markup does:
 * `rebuildCellTextBody` collapses the cell to a SINGLE run carrying the first
 * run's `a:rPr`, and `PptxTableCell.style` is that same first-run style, so the
 * flat-text fallback and the rewritten `a:txBody` paint the same thing.
 * For the same reason only the first paragraph's layout is kept: the
 * rewritten `a:txBody` keeps only the first paragraph's `a:pPr`.
 *
 * @param cell - The cell to re-text (not mutated).
 * @param text - New plain-text content.
 * @returns A new cell holding `text` and no stale run model.
 */
export function withCellText(cell: PptxTableCell, text: string): PptxTableCell {
	const next: PptxTableCell = { ...cell, text };
	delete next.textRuns;
	if (cell.paragraphs) {
		next.paragraphs = cell.paragraphs.slice(0, 1);
	}
	return next;
}

/**
 * Merge a style patch into one cell.
 *
 * Aligning a cell aligns every paragraph in it: save writes each paragraph's
 * own alignment, so a paragraph that kept its old one would keep it in the
 * file. `paragraphs` gets one entry per line of `text`; save writes as many as
 * the cell has `a:p`.
 *
 * @param cell - The cell to restyle (not mutated).
 * @param patch - Style fields to set.
 * @returns A new cell with the patch applied.
 */
export function withCellStyle(
	cell: PptxTableCell,
	patch: Partial<PptxTableCellStyle>,
): PptxTableCell {
	const next: PptxTableCell = { ...cell, style: { ...cell.style, ...patch } };
	const { align } = patch;
	if (align) {
		next.paragraphs = cell.text
			.split('\n')
			.map((_, index) => ({ ...cell.paragraphs?.[index], align }));
	}
	return next;
}

/**
 * Return a new `TablePptxElement` with the text of a single cell replaced.
 *
 * The element is not mutated: the affected row and cell are shallow-cloned and
 * every other row/cell is reused by reference. Returns the original element
 * unchanged when it carries no `tableData`.
 *
 * @param element - The source table element (not mutated).
 * @param rowIndex - Zero-based row index of the cell.
 * @param colIndex - Zero-based column index of the cell.
 * @param text - New plain-text content for the cell.
 * @returns A new `TablePptxElement` with the cell text applied.
 *
 * @example
 * ```ts
 * const updated = setCellText(el, 0, 1, "Revenue");
 * ```
 */
export function setCellText(
	element: TablePptxElement,
	rowIndex: number,
	colIndex: number,
	text: string,
): TablePptxElement {
	const tableData = element.tableData;
	if (!tableData) {
		return element;
	}
	const rows = tableData.rows.map((row, ri) => {
		if (ri !== rowIndex) {
			return row;
		}
		return {
			...row,
			cells: row.cells.map((cell, ci) => (ci === colIndex ? withCellText(cell, text) : cell)),
		};
	});
	return { ...element, tableData: { ...tableData, rows } };
}
