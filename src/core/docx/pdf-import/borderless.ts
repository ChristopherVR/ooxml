/**
 * Borderless tables: a line cut by column-gap tabs into two or more pieces starts a run, and the
 * lines after it join while their text keeps to the same columns (the whitespace corridors between
 * the columns stay empty). Cells whose text wraps and empty cells are allowed: the run is cut into
 * rows where a cell's text starts again rather than going on from a full line above.
 */
import type { Fragment, Line } from './lines';
import { firstWordWidth, inkRight, splitLine, subLine } from './sublines';
import { cellParagraphs, DEFAULT_PADDING, type TableBlock } from './tables';

interface Column {
	id: number;
	left: number;
	right: number;
	/** The pieces placed in the column so far (alignment evidence). */
	pieces: Line[];
}

interface Entry {
	line: Line;
	/** The line's text in each column (undefined where the cell has none on this line). */
	parts: (Line | undefined)[];
}

const blank = (fragment: Fragment) => !fragment.text.trim();

/** A line's pieces between column-gap tabs (none when the line has no such tab). */
function tabPieces(line: Line): Line[] {
	const cuts: number[] = [];
	line.fragments.forEach((fragment, index) => {
		const next = line.fragments.slice(index + 1).find((item) => item.text.trim());
		// A leader tab (a table of contents) is not a column gap.
		if (/\t\s*$/.test(fragment.text) && next && !fragment.leader) cuts.push(next.x - 0.5);
	});
	if (!cuts.length) return [];
	const parts = splitLine(line, cuts);
	return parts.some((part) => !part) ? [] : (parts as Line[]);
}

const aligned = (a: { left: number; right: number }, b: { left: number; right: number }) =>
	Math.abs(a.left - b.left) < 3 ||
	Math.abs(a.right - b.right) < 3 ||
	Math.abs(a.left + a.right - b.left - b.right) < 6;

const overlap = (left: number, right: number, column: Column) =>
	Math.min(right, column.right) - Math.max(left, column.left);

/** Puts a line's fragments in the columns, or returns undefined when it does not fit them. */
function place(
	line: Line,
	columns: Column[],
): { parts: (Line | undefined)[]; columns: Column[] } | undefined {
	const minGap = 0.5 * line.size;
	const cols = columns.map((column) => ({ ...column }));
	// A column-gap piece outside every column starts one (an empty top-left header cell).
	const tabs = tabPieces(line);
	if (tabs.length >= 2)
		for (const piece of tabs)
			if (!cols.some((column) => overlap(piece.left, piece.right, column) > 0))
				cols.push({
					id: Math.max(...cols.map((column) => column.id)) + 1,
					left: piece.left,
					right: piece.right,
					pieces: [],
				});
	cols.sort((a, b) => a.left - b.left);
	const groups: Fragment[][] = cols.map(() => []);
	let current = -1;
	for (const fragment of line.fragments) {
		if (!blank(fragment)) {
			const [left, right] = [fragment.x, inkRight(fragment)];
			let best = -1;
			let score = -Infinity;
			cols.forEach((column, at) => {
				// Overlap first; a fragment beside every column goes to the nearest.
				const value = overlap(left, right, column);
				if (value > score) [best, score] = [at, value];
			});
			current = best;
			const column = cols[best]!;
			column.left = Math.min(column.left, left);
			column.right = Math.max(column.right, right);
		}
		if (current >= 0) groups[current]!.push(fragment);
	}
	for (let at = 1; at < cols.length; at++)
		if (cols[at]!.left - cols[at - 1]!.right < minGap) return undefined;
	const parts = groups.map((fragments) => subLine(line, fragments));
	const filled = parts.filter(Boolean).length;
	if (!filled) return undefined;
	// A lone piece (a wrapped line) keeps to an edge or the middle of text already in its column.
	if (
		filled === 1 &&
		!parts.every((part, at) => !part || cols[at]!.pieces.some((piece) => aligned(part, piece)))
	)
		return undefined;
	parts.forEach((part, at) => {
		if (part) cols[at]!.pieces = [...cols[at]!.pieces, part];
	});
	return { parts, columns: cols };
}

/** The run of lines from `start` that keeps to one set of columns. */
function collect(lines: Line[], start: number): { entries: Entry[]; columns: Column[] } {
	const first = lines[start]!;
	const seed = tabPieces(first);
	if (seed.length < 2) return { entries: [], columns: [] };
	let columns: Column[] = seed.map((piece, id) => ({
		id,
		left: piece.left,
		right: piece.right,
		pieces: [piece],
	}));
	// Parts by column id, since a later line can add a column before the others.
	const byId: { line: Line; parts: Map<number, Line> }[] = [
		{ line: first, parts: new Map(seed.map((piece, id) => [id, piece])) },
	];
	const pitches: number[] = [];
	for (let at = start + 1; at < lines.length; at++) {
		const line = lines[at]!;
		const pitch = line.baseline - lines[at - 1]!.baseline;
		if (
			pitch <= 0 ||
			pitch > 2.5 * line.size ||
			Math.abs(line.size - first.size) > 0.25 * first.size
		)
			break;
		const placed = place(line, columns);
		if (!placed) break;
		const base = pitches.length ? Math.min(...pitches) : 1.25 * first.size;
		if (placed.parts.filter(Boolean).length === 1 && pitch > 1.35 * base) break;
		columns = placed.columns;
		const parts = new Map<number, Line>();
		placed.parts.forEach((part, col) => {
			if (part) parts.set(columns[col]!.id, part);
		});
		byId.push({ line, parts });
		pitches.push(pitch);
	}
	const entries = byId.map(({ line, parts }) => ({
		line,
		parts: columns.map((column) => parts.get(column.id)),
	}));
	return { entries, columns };
}

/** Cuts the run into rows: lines whose every piece goes on from a full line above stay in the row. */
function rowsOf(entries: Entry[], columns: Column[]): Entry[][] {
	const pitches = entries
		.slice(1)
		.map((entry, at) => entry.line.baseline - entries[at]!.line.baseline);
	const base = pitches.length ? Math.min(...pitches) : 0;
	const limit = (col: number) =>
		col + 1 < columns.length
			? Math.max(columns[col]!.right, columns[col + 1]!.left - 2 * DEFAULT_PADDING)
			: columns[col]!.right;
	const rows: Entry[][] = [];
	entries.forEach((entry, at) => {
		const previous = entries[at - 1];
		const row = rows[rows.length - 1];
		const goesOn =
			previous &&
			row &&
			pitches[at - 1]! <= base * 1.2 + 1 &&
			entry.parts.every((part, col) => {
				if (!part) return true;
				const above = previous.parts[col];
				if (!above) return false;
				const wrapped = above.right + 0.25 * part.size + firstWordWidth(part) > limit(col) - 1;
				// A further list item in a cell whose items started in this row.
				const item =
					/^\s*(?:[\u2022\u25cf\u25aa\u25a0\u00b7\u2219\uf000-\uf0ff-]|\(?\w{1,4}[.)])\s/.test(
						part.text,
					) && Math.abs(part.left - above.left) < 3;
				return wrapped || item;
			});
		if (goesOn) row.push(entry);
		else rows.push([entry]);
	});
	return rows;
}

const filledCells = (row: Entry[], count: number) =>
	Array.from({ length: count }, (_, col) => row.some((entry) => entry.parts[col])).filter(Boolean)
		.length;

const MARKER_ONLY =
	/^(?:\(?(?:\d{1,3}|[a-zA-Z]|[ivxlcdmIVXLCDM]{1,6})[.)]|[\u2022\u25cf\u25aa\u25a0\u00b7\u2219\u2013\u2014\-o*\uf000-\uf0ff])$/;

/** Whether the rows make a table rather than lists or paragraphs with tabs. */
function plausible(rows: Entry[][], count: number): boolean {
	if (rows.length < 3 || count < 2) return false;
	if (rows.filter((row) => filledCells(row, count) >= 2).length * 2 < rows.length) return false;
	for (let col = 0; col < count; col++)
		if (rows.filter((row) => row.some((entry) => entry.parts[col])).length < 2) return false;
	// A first column of only list markers ("1.", "-") is a list set off with tabs.
	const markers = rows.flatMap((row) =>
		row.flatMap((entry) => (entry.parts[0] ? [entry.parts[0].text.trim()] : [])),
	);
	return !markers.every((text) => MARKER_ONLY.test(text));
}

/** Borderless tables: runs of lines whose text keeps to the same columns. */
export function unruledTables(
	page: number,
	lines: Line[],
	bodySize: number,
): { tables: TableBlock[]; lines: Line[] } {
	const tables: TableBlock[] = [];
	const rest: Line[] = [];
	let index = 0;
	while (index < lines.length) {
		const { entries, columns } = collect(lines, index);
		const rows = entries.length ? rowsOf(entries, columns) : [];
		// A last row with one cell is more likely the paragraph after the table.
		while (rows.length && filledCells(rows[rows.length - 1]!, columns.length) < 2) rows.pop();
		if (!plausible(rows, columns.length)) {
			rest.push(lines[index]!);
			index++;
			continue;
		}
		tables.push(borderless(page, rows, columns, bodySize));
		index += rows.reduce((sum, row) => sum + row.length, 0);
	}
	return { tables, lines: rest };
}

function borderless(
	page: number,
	rows: Entry[][],
	columns: Column[],
	bodySize: number,
): TableBlock {
	const count = columns.length;
	const starts = columns.map((column) => column.left);
	// The last column gets some slack, so its longest text does not wrap in Word's own metrics.
	const size = rows[0]![0]!.line.size;
	const end = columns[count - 1]!.right + size;
	const edges = [...starts.map((start) => start - DEFAULT_PADDING), end + DEFAULT_PADDING];
	const firsts = rows.map((row) => row[0]!.line);
	const last = rows[rows.length - 1]!;
	const top = firsts[0]!.top - 1;
	const bottom = last[last.length - 1]!.line.bottom + 1;
	return {
		kind: 'table',
		page,
		top,
		bottom,
		left: edges[0]!,
		columns: edges.slice(1).map((edge, at) => edge - edges[at]!),
		rows: firsts.map((line, at) => ({
			height: (firsts[at + 1]?.top ?? bottom) - (at ? line.top : top),
		})),
		padding: DEFAULT_PADDING,
		cells: rows.flatMap((row, r) =>
			columns.map((_, col) => ({
				row: r,
				col,
				rowSpan: 1,
				colSpan: 1,
				borders: {},
				blocks: cellParagraphs(
					row.flatMap((entry) => (entry.parts[col] ? [entry.parts[col]] : [])),
					edges[col]! + DEFAULT_PADDING,
					edges[col + 1]! - DEFAULT_PADDING,
					bodySize,
				),
			})),
		),
	};
}
