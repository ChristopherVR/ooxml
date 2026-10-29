// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Merging and splitting cells on a table with merged cells. Results follow Word's
// (src/core/docx/__fixtures__/word-table-structure.json).
import type { Paragraph, Table, TableCell } from './model';
import {
	continuationsOf,
	mergeOrigin,
	rowSpanOf,
	tableGrid,
	type GridCell,
	type TableGrid,
} from './table-grid';
import {
	copiedRowProperties,
	copyTable,
	emptyCell,
	paddedRowProperties,
	storeRowProperties,
	type IdFactory,
} from './table-structure-cells';
import { normalizeMerges } from './table-structure';
import type { Twips } from './units';

/** A rectangle of grid positions, inclusive on all four sides. */
export interface GridRect {
	firstRow: number;
	lastRow: number;
	firstCol: number;
	lastCol: number;
}

const isEmpty = (paragraph: Paragraph): boolean =>
	paragraph.runs.every((run) => !run.text && !run.contentControl);

/** Whether `cell` holds anything but empty paragraphs (text, a nested table, a content control). */
export const hasContent = (cell: TableCell): boolean =>
	!(cell.paragraphs.length === 1 && cell.paragraphs.every(isEmpty)) ||
	Boolean(cell.sourceBlocks?.some((block) => block.type === 'table')) ||
	Boolean(cell.nestedTables?.length);

const sumWidths = (table: Table, from: number, to: number): Twips | undefined =>
	table.grid ? (table.grid.slice(from, to + 1).reduce((sum, w) => sum + w, 0) as Twips) : undefined;

/** The distinct cells inside `rect`, or undefined when a merged cell sticks out of it. */
function cellsWithin(grid: TableGrid, rect: GridRect): GridCell[] | undefined {
	const found = new Set<GridCell>();
	for (let r = rect.firstRow; r <= rect.lastRow; r++)
		for (let c = rect.firstCol; c <= rect.lastCol; c++) {
			const entry = grid.cover[r]?.[c];
			if (!entry) return undefined;
			found.add(entry);
		}
	for (const entry of found) {
		if (entry.col < rect.firstCol || entry.col + entry.span - 1 > rect.lastCol) return undefined;
		const origin = entry.cell.verticalMerge === 'continue' ? undefined : entry;
		if (origin) {
			if (origin.row + rowSpanOf(grid, origin) - 1 > rect.lastRow) return undefined;
		} else if (!found.has(grid.cover[entry.row - 1]?.[entry.col] as GridCell)) return undefined;
	}
	return [...found];
}

/** Grows `rect` until no cell is cut by its edge: every merged cell it touches lies wholly inside. */
export function closeRect(table: Table, rect: GridRect): GridRect | undefined {
	const grid = tableGrid(table);
	if (!grid) return undefined;
	let current = { ...rect };
	for (let pass = 0; pass <= grid.columns + grid.rows.length; pass++) {
		let next = current;
		for (let r = current.firstRow; r <= current.lastRow; r++)
			for (let c = current.firstCol; c <= current.lastCol; c++) {
				const entry = grid.cover[r]?.[c];
				if (!entry) return undefined;
				const origin = mergeOrigin(grid, entry);
				next = {
					firstRow: Math.min(next.firstRow, origin.row),
					lastRow: Math.max(next.lastRow, origin.row + rowSpanOf(grid, origin) - 1),
					firstCol: Math.min(next.firstCol, origin.col),
					lastCol: Math.max(next.lastCol, origin.col + origin.span - 1),
				};
			}
		if (JSON.stringify(next) === JSON.stringify(current)) return current;
		current = next;
	}
	return current;
}

/**
 * Merges the cells in `rect` into one: the first cell keeps its formatting and takes the others'
 * non-empty paragraphs after its own. Across rows the result is a vertical merge of
 * `lastCol - firstCol + 1` columns.
 */
export function mergeCells(table: Table, rect: GridRect, newId: IdFactory): Table | undefined {
	const grid = tableGrid(table);
	if (!grid) return undefined;
	const within = cellsWithin(grid, rect);
	if (!within) return undefined;
	const origins = within
		.filter((entry) => entry.cell.verticalMerge !== 'continue')
		.sort((a, b) => a.row - b.row || a.col - b.col);
	const [head, ...absorbed] = origins;
	if (!head || !absorbed.length) return undefined;
	if (absorbed.some((entry) => entry.cell.sourceBlocks?.length || entry.cell.nestedTables?.length))
		return undefined;
	const next = copyTable(table);
	const keep = next.rows[head.row]?.[head.index];
	if (!keep) return undefined;
	const paragraphs = [...keep.paragraphs];
	const extra = absorbed.flatMap((entry) =>
		hasContent(entry.cell) ? entry.cell.paragraphs.filter((p) => !isEmpty(p)) : [],
	);
	if (extra.length) {
		const lead = paragraphs.length === 1 && paragraphs.every(isEmpty) && !hasContent(keep);
		keep.paragraphs = lead ? structuredClone(extra) : [...paragraphs, ...structuredClone(extra)];
		if (lead) keep.paragraphs = [...structuredClone(extra)];
	}
	const width = rect.lastCol - rect.firstCol + 1;
	if (width > 1) keep.gridSpan = width;
	else delete keep.gridSpan;
	const total = sumWidths(table, rect.firstCol, rect.lastCol);
	if (keep.widthTwips !== undefined && total !== undefined) keep.widthTwips = total;
	if (rect.lastRow > rect.firstRow) keep.verticalMerge = 'restart';
	else delete keep.verticalMerge;
	for (let r = rect.firstRow; r <= rect.lastRow; r++) {
		const row = next.rows[r];
		if (!row) continue;
		const source = grid.rows[r] ?? [];
		const inside = source.filter(
			(entry) => entry.col >= rect.firstCol && entry.col + entry.span - 1 <= rect.lastCol,
		);
		const index = inside[0]?.index ?? 0;
		const replacement =
			r === rect.firstRow
				? keep
				: Object.assign(emptyCell(newId, keep, width, keep.widthTwips), {
						verticalMerge: 'continue' as const,
					});
		row.splice(index, inside.length, replacement);
	}
	normalizeMerges(next);
	return next;
}

/** Spans of `parts` sub-cells covering `span` columns, as even as the columns allow. */
function distribute(span: number, parts: number): number[] {
	const base = Math.floor(span / parts);
	const rest = span % parts;
	return Array.from({ length: parts }, (_, i) => base + (i < rest ? 1 : 0));
}

/**
 * Splits the cell at `row`, `index` into `columns` side-by-side cells and `rows` stacked ones.
 * The text stays in the first cell. A cell of one column split into several adds grid columns,
 * which the cells above and below widen to cover; a cell split into rows makes the other cells
 * of its row vertical merges that run through the new rows. A vertical merge splits into
 * `rows` equal to its height (it comes apart) or 1.
 */
export function splitCell(
	table: Table,
	row: number,
	index: number,
	rows: number,
	columns: number,
	newId: IdFactory,
): Table | undefined {
	const grid = tableGrid(table);
	const origin = grid?.rows[row]?.[index];
	if (!grid || !origin || origin.cell.verticalMerge === 'continue') return undefined;
	if (rows < 1 || columns < 1 || (rows === 1 && columns === 1)) return undefined;
	const height = rowSpanOf(grid, origin);
	if (height > 1 && rows !== height && !(rows === 1 && columns === 1)) return undefined;
	// Word re-grids a merged cell split into other than its own columns; only the clean cases are done.
	if (origin.span > 1 && columns !== origin.span && columns !== 1) return undefined;
	if (height > 1 && rows === 1 && columns > 1) return undefined;
	const next = copyTable(table);
	const members = [origin, ...continuationsOf(grid, origin)];
	const widths = table.grid ? table.grid.slice(origin.col, origin.col + origin.span) : [];
	const total = widths.reduce((sum, w) => sum + w, 0) as Twips;
	const grow = columns > origin.span;
	const spans = grow ? Array.from({ length: columns }, () => 1) : distribute(origin.span, columns);
	const partWidth = (i: number): Twips | undefined => {
		if (origin.cell.widthTwips === undefined && !total) return undefined;
		if (grow) return Math.floor(total / columns) as Twips;
		let from = origin.col + spans.slice(0, i).reduce((s, v) => s + v, 0);
		return sumWidths(table, from, from + (spans[i] ?? 1) - 1);
	};
	if (grow && next.grid) {
		const each = Array.from({ length: columns }, (_, i) =>
			i === columns - 1
				? ((total - Math.floor(total / columns) * (columns - 1)) as Twips)
				: (Math.floor(total / columns) as Twips),
		);
		next.grid.splice(origin.col, 1, ...each);
	}
	// Other rows: any cell over the split column grows to keep covering it.
	if (grow)
		grid.rows.forEach((cells, r) => {
			if (members.some((m) => m.row === r)) return;
			const over = cells.find((e) => e.col <= origin.col && origin.col < e.col + e.span);
			const cell = over && next.rows[r]?.[over.index];
			if (cell) cell.gridSpan = (cell.gridSpan ?? 1) + columns - 1;
		});
	// Column split of every member (restart and its continuations stay vertically merged).
	[...members].reverse().forEach((member) => {
		const target = next.rows[member.row];
		const cell = target?.[member.index];
		if (!target || !cell) return;
		const parts: TableCell[] = spans.map((span, i) => {
			if (i === 0) return cell;
			const part = emptyCell(newId, cell, span, partWidth(i));
			if (cell.verticalMerge) part.verticalMerge = cell.verticalMerge;
			return part;
		});
		// Word hands a cell's paragraphs out one to each new cell in a plain column split (the last
		// takes any that remain); in a split into rows or a grid the first cell keeps them all.
		const own = member === origin ? [...cell.paragraphs] : [];
		if (rows === 1 && own.length > 1)
			parts.forEach((part, i) => {
				const mine = i === parts.length - 1 ? own.slice(i) : own.slice(i, i + 1);
				if (mine.length) part.paragraphs = mine;
			});
		const first = parts[0];
		if (first) {
			if ((spans[0] ?? 1) > 1) first.gridSpan = spans[0] ?? 1;
			else delete first.gridSpan;
			const w = partWidth(0);
			if (w !== undefined && first.widthTwips !== undefined) first.widthTwips = w;
		}
		target.splice(member.index, 1, ...parts);
	});
	if (rows > 1) {
		if (height === 1) stackRows(next, grid, origin, rows, columns, newId);
		else for (const member of members) uncouple(next, member, columns);
	}
	normalizeMerges(next);
	return next;
}

/** Splitting a vertical merge into its rows: every cell of it stands alone. */
function uncouple(next: Table, member: GridCell, columns: number): void {
	const row = next.rows[member.row];
	if (!row) return;
	for (let k = 0; k < columns; k++) delete row[member.index + k]?.verticalMerge;
}

/** Adds `rows - 1` rows after the split cell's row and ties the other cells to them. */
function stackRows(
	next: Table,
	grid: TableGrid,
	origin: GridCell,
	rows: number,
	columns: number,
	newId: IdFactory,
): void {
	const source = next.rows[origin.row] ?? [];
	const splitStart = origin.index;
	const properties = paddedRowProperties(next);
	for (let step = 1; step < rows; step++) {
		const added = source.map((cell, i) => {
			const own = i >= splitStart && i < splitStart + columns;
			const copy = emptyCell(newId, cell, cell.gridSpan ?? 1, cell.widthTwips);
			if (!own) copy.verticalMerge = 'continue';
			return copy;
		});
		next.rows.splice(origin.row + step, 0, added);
		properties.splice(origin.row + step, 0, copiedRowProperties(next, origin.row));
	}
	source.forEach((cell, i) => {
		if (i >= splitStart && i < splitStart + columns) return;
		if (!cell.verticalMerge) cell.verticalMerge = 'restart';
	});
	// A merge that already ran through the row, or continues past it, now runs through the new rows.
	void grid;
	storeRowProperties(next, properties);
}
