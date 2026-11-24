/**
 * Word tables from the page: the text of a ruled grid goes to its cells (several paragraphs per
 * cell); borderless tables come from `borderless.ts`. A table that goes on at the top of the next
 * page with the same columns joins the one before, dropping the header row the PDF repeats and
 * joining a row the page break cut (`row-split.ts`).
 */
import type { Line } from './lines';
import { buildParagraphs, type ParagraphBlock } from './paragraphs';
import { joinSplitRow } from './row-split';
import { splitLine } from './sublines';
import type { GridBorder, RuledGrid } from './table-grid';

export interface TableCellBlock {
	row: number;
	col: number;
	rowSpan: number;
	colSpan: number;
	fill?: string;
	/** Unset sides have no border. */
	borders: { top?: GridBorder; left?: GridBorder; bottom?: GridBorder; right?: GridBorder };
	blocks: ParagraphBlock[];
}

export interface TableBlock {
	kind: 'table';
	page: number;
	top: number;
	bottom: number;
	left: number;
	/** Column widths, points. */
	columns: number[];
	rows: { height: number; header?: boolean }[];
	cells: TableCellBlock[];
	/** Left and right cell padding, points. */
	padding: number;
	/** Distance from the text area's left edge to the table's, points (set by the page flow). */
	indent?: number;
}

export const DEFAULT_PADDING = 5.4;

/** A cell's paragraphs, between its padded edges. */
export function cellParagraphs(lines: Line[], left: number, right: number, bodySize: number) {
	return buildParagraphs(lines, { left, right }, bodySize, { wrapped: true }).map(
		({ heading: _heading, ...block }, index) => ({
			...block,
			...(index ? {} : { spaceBefore: 0 }),
		}),
	);
}

const median = (values: number[], fallback: number) => {
	if (!values.length) return fallback;
	const sorted = [...values].sort((a, b) => a - b);
	return sorted[Math.floor(sorted.length / 2)]!;
};

/** Takes the text inside each grid out of `lines` and returns the tables, top to bottom. */
export function ruledTables(
	page: number,
	lines: Line[],
	grids: RuledGrid[],
	bodySize: number,
): { tables: TableBlock[]; lines: Line[] } {
	let rest = lines;
	const tables: TableBlock[] = [];
	for (const grid of grids) {
		const { xs, ys } = grid;
		const [x0, x1, y0, y1] = [xs[0]!, xs[xs.length - 1]!, ys[0]!, ys[ys.length - 1]!];
		const byCell = new Map<number, Line[]>();
		const kept: Line[] = [];
		let used = false;
		for (const line of rest) {
			const middle = line.baseline - line.size * 0.3;
			if (middle < y0 || middle > y1 || line.right < x0 || line.left > x1) {
				kept.push(line);
				continue;
			}
			const below = ys.findIndex((y, index) => index > 0 && middle < y);
			const row = below < 0 ? ys.length - 2 : below - 1;
			// Pieces left and right of the table stay outside it; the rest split at the column edges.
			const pieces = splitLine(line, xs);
			if (pieces[0]) kept.push(pieces[0]);
			if (pieces[xs.length]) kept.push(pieces[xs.length]!);
			for (let col = 0; col < xs.length - 1; col++) {
				const piece = pieces[col + 1];
				if (!piece) continue;
				const cell = grid.cells.find(
					(item) =>
						row >= item.row &&
						row < item.row + item.rowSpan &&
						col >= item.col &&
						col < item.col + item.colSpan,
				);
				if (!cell) continue;
				const key = grid.cells.indexOf(cell);
				byCell.set(key, [...(byCell.get(key) ?? []), piece]);
				used = true;
			}
		}
		if (!used) continue;
		rest = kept;
		const padding = Math.min(
			12,
			Math.max(
				0,
				median(
					[...byCell].map(
						([key, cellLines]) =>
							Math.min(...cellLines.map((line) => line.left)) - xs[grid.cells[key]!.col]!,
					),
					DEFAULT_PADDING,
				),
			),
		);
		tables.push({
			kind: 'table',
			page,
			top: y0,
			bottom: y1,
			left: x0,
			columns: xs.slice(1).map((x, index) => x - xs[index]!),
			rows: ys.slice(1).map((y, index) => ({ height: y - ys[index]! })),
			padding,
			cells: grid.cells.map((cell, key) => ({
				row: cell.row,
				col: cell.col,
				rowSpan: cell.rowSpan,
				colSpan: cell.colSpan,
				...(cell.fill ? { fill: cell.fill } : {}),
				borders: cell.borders,
				blocks: cellParagraphs(
					byCell.get(key) ?? [],
					xs[cell.col]! + padding,
					xs[cell.col + cell.colSpan]! - padding,
					bodySize,
				),
			})),
		});
	}
	return { tables, lines: rest };
}

const rowText = (table: TableBlock, row: number) =>
	table.cells
		.filter((cell) => cell.row === row)
		.map((cell) =>
			cell.blocks.map((block) => block.fragments.map((fragment) => fragment.text).join('')),
		)
		.join('|');

/** Joins `next` (at the top of a page) to `table` (ending the page before) when it goes on. */
export function joinTables(table: TableBlock, next: TableBlock): boolean {
	if (
		table.columns.length !== next.columns.length ||
		Math.abs(table.left - next.left) > 3 ||
		table.columns.some((width, at) => Math.abs(width - next.columns[at]!) > 3)
	)
		return false;
	let skip = 0;
	if (next.rows.length > 1 && rowText(table, 0) === rowText(next, 0) && rowText(next, 0)) {
		table.rows[0]!.header = true;
		skip = 1;
	}
	// The first row the page shows may be the rest of the row the page break cut.
	if (skip < next.rows.length && joinSplitRow(table, next, skip)) skip++;
	const offset = table.rows.length - skip;
	table.rows.push(...next.rows.slice(skip));
	for (const cell of next.cells)
		if (cell.row >= skip) table.cells.push({ ...cell, row: cell.row + offset });
	return true;
}
