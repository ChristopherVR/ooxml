/**
 * A table row that a page break cut: Word draws each part as a row of its own (with its own
 * borders), so the row a page starts with is joined to the last row of the page before when it
 * leaves empty the cells that are never empty elsewhere in the table, or when every cell with text
 * goes on mid-sentence.
 */
import { continuesOnNextPage, type ParagraphBlock } from './paragraphs';
import type { TableBlock, TableCellBlock } from './tables';

const hasText = (cell: TableCellBlock | undefined) =>
	!!cell?.blocks.some((block) => block.fragments.some((fragment) => fragment.text.trim()));

/** Whether a cell's last paragraph goes on in the first paragraph of `after`. */
function midParagraph(before: TableCellBlock, after: TableCellBlock): boolean {
	const last = before.blocks[before.blocks.length - 1];
	const first = after.blocks[0];
	if (!last || !first || !continuesOnNextPage(last, first)) return false;
	return /^\s*[a-z\u00df-\u00ff]/.test(first.lines[0]!.text);
}

function joinParagraph(last: ParagraphBlock, first: ParagraphBlock): void {
	const tail = last.fragments[last.fragments.length - 1];
	if (tail) {
		const hyphen = /[-\u00ad]$/.test(tail.text);
		tail.text = hyphen ? tail.text.replace(/[-\u00ad]$/, '') : `${tail.text} `;
	}
	last.fragments.push(...first.fragments);
	last.lines.push(...first.lines);
}

const rowCells = (table: TableBlock, row: number) =>
	table.cells.filter((cell) => cell.row === row).sort((a, b) => a.col - b.col);

/**
 * Joins row `row` of `next` (a table at the top of a page) to the last row of `table` when it is
 * the rest of that row; returns whether it did.
 */
export function joinSplitRow(table: TableBlock, next: TableBlock, row: number): boolean {
	const last = table.rows.length - 1;
	const before = rowCells(table, last);
	const after = rowCells(next, row);
	if (!before.length || before.length !== after.length) return false;
	// Only rows with the same cells and no vertical merge reaching into them.
	if (
		before.some(
			(cell, at) =>
				cell.rowSpan > 1 ||
				after[at]!.rowSpan > 1 ||
				cell.col !== after[at]!.col ||
				cell.colSpan !== after[at]!.colSpan,
		) ||
		table.cells.some((cell) => cell.row < last && cell.row + cell.rowSpan > last)
	)
		return false;
	if (!after.some(hasText)) return false;
	// The body rows around the cut, to tell a column that is never empty from one that often is.
	const others = [
		...table.cells.filter((cell) => cell.row < last && !table.rows[cell.row]?.header),
		...next.cells.filter((cell) => cell.row > row),
	];
	const neverEmpty = (cell: TableCellBlock) => {
		const column = others.filter(
			(other) => other.col === cell.col && other.colSpan === cell.colSpan,
		);
		return column.length >= 2 && column.every(hasText);
	};
	const emptied =
		before.some((cell, at) => hasText(cell) && !hasText(after[at]) && neverEmpty(cell)) &&
		before.every((cell, at) => hasText(cell) || !hasText(after[at]));
	const continued = before.every(
		(cell, at) => !hasText(after[at]) || (hasText(cell) && midParagraph(cell, after[at]!)),
	);
	if (!emptied && !continued) return false;
	before.forEach((cell, at) => {
		const rest = after[at]!;
		const blocks = [...rest.blocks];
		if (blocks.length && midParagraph(cell, rest))
			joinParagraph(cell.blocks[cell.blocks.length - 1]!, blocks.shift()!);
		cell.blocks.push(...blocks);
		if (rest.borders.bottom) cell.borders.bottom = rest.borders.bottom;
		else delete cell.borders.bottom;
	});
	table.rows[last]!.height += next.rows[row]!.height;
	return true;
}
