import {
	type CellAddress,
	type CellRange,
	MAX_COL,
	MAX_ROW,
	normalizeRange,
	rangeContains,
	rangesIntersect,
} from '../address.js';
import { deleteCell, forEachCellInRange, getCell, putCell, usedRange } from '../cells.js';
import type { Cell, Workbook } from '../model.js';
import { internStyle, styleAt } from '../styles.js';
import { clearContents } from './cell-values.js';
import { toHtml } from './clipboard-html.js';
import { parseHtmlTable } from './clipboard-html-parse.js';
import { parseTsv, toTsv } from './clipboard-text.js';
import { type EditContext, displayText, ensureCell, pruneCell, sheetAt } from './context.js';
import { isSpilledCell, moveReferencesInFormula, parseCellInput } from './deps.js';
import { rewriteFormulas } from './shift-formulas.js';
import { moveFormula } from './fill.js';
import type { EditScope } from './history.js';
import { rangeWithin } from './range-math.js';
import type { ClipboardCell, ClipboardCells, ClipboardPayload, PasteMode } from './types.js';

/** Copies a range into a self-contained payload (whole rows or columns stop at the used area). */
export function copyRange(workbook: Workbook, s: number, range: CellRange): ClipboardPayload {
	const sheet = sheetAt(workbook, s);
	const r = normalizeRange(range);
	// The used area, grown to the merges inside the range (a merge's covered cells are empty).
	const used = usedRange(sheet);
	let lastRow = used?.end.row ?? r.start.row;
	let lastCol = used?.end.col ?? r.start.col;
	for (const m of sheet.merges)
		if (rangeWithin(m, r)) {
			lastRow = Math.max(lastRow, m.end.row);
			lastCol = Math.max(lastCol, m.end.col);
		}
	const end = {
		row: Math.max(r.start.row, Math.min(r.end.row, lastRow)),
		col: Math.max(r.start.col, Math.min(r.end.col, lastCol)),
	};
	const clipped: CellRange = { start: r.start, end };
	const rows = end.row - r.start.row + 1;
	const cols = end.col - r.start.col + 1;
	const data: (ClipboardCell | null)[][] = Array.from({ length: rows }, () =>
		Array.from({ length: cols }, () => null),
	);
	forEachCellInRange(sheet, clipped, (cell, row, col) => {
		const clip: ClipboardCell = {
			value: cell.value,
			style: structuredClone(styleAt(workbook, cell.styleId)),
			text: displayText(workbook, cell),
		};
		if (cell.formula !== undefined) clip.formula = cell.formula;
		if (cell.legacyFormula) clip.legacyFormula = true;
		if (cell.richText) clip.richText = structuredClone(cell.richText);
		if (isSpilledCell(cell) && rangeContains(clipped, cell.spillAnchor)) clip.spilled = true;
		const line = data[row - r.start.row];
		if (line) line[col - r.start.col] = clip;
	});
	const merges = sheet.merges
		.filter((m) => rangeWithin(m, clipped))
		.map((m) => ({
			start: { row: m.start.row - r.start.row, col: m.start.col - r.start.col },
			end: { row: m.end.row - r.start.row, col: m.end.col - r.start.col },
		}));
	const cells: ClipboardCells = { rows, cols, data, merges, source: { sheet: s, range: clipped } };
	return {
		tsv: toTsv(data.map((line) => line.map((c) => c?.text ?? ''))),
		html: toHtml(cells, workbook.theme),
		cells,
	};
}

/** Clipboard text (HTML table or TSV) as cells; plain text goes through input parsing. */
export function cellsFromText(workbook: Workbook, text: string): ClipboardCells {
	if (/<table\b/i.test(text)) {
		const parsed = parseHtmlTable(text, styleAt(workbook, 0));
		if (parsed) return parsed;
	}
	const grid = parseTsv(text);
	const data = grid.map((line) =>
		line.map((field): ClipboardCell | null => {
			if (field === '') return null;
			const parsed = parseCellInput(field, { date1904: workbook.date1904 });
			const clip: ClipboardCell = { value: parsed.value, text: field };
			if (parsed.formula !== undefined) clip.formula = parsed.formula;
			if (parsed.numFmt) clip.numFmt = parsed.numFmt;
			return clip;
		}),
	);
	return { rows: grid.length, cols: grid[0]?.length ?? 0, data, merges: [] };
}

/**
 * Pastes clipboard cells at `at`. `all` pastes values, formulas (moved relative to the copy) and
 * formats; `values` the results only; `formulas` formulas without formats; `formats` formats only;
 * `transpose` everything with rows and columns swapped. A cut payload clears its source on the
 * first paste and keeps formulas unchanged. Returns the pasted range.
 */
export function pasteAt(
	ctx: EditContext,
	s: number,
	at: CellAddress,
	payload: ClipboardPayload | string,
	mode: PasteMode,
): CellRange {
	const { workbook } = ctx;
	const sheet = sheetAt(workbook, s);
	const cells = typeof payload === 'string' ? cellsFromText(workbook, payload) : payload.cells;
	const transpose = mode === 'transpose';
	const height = transpose ? cells.cols : cells.rows;
	const width = transpose ? cells.rows : cells.cols;
	if (!height || !width) return { start: at, end: at };
	if (at.row + height - 1 > MAX_ROW || at.col + width - 1 > MAX_COL)
		throw new RangeError('The paste area extends beyond the sheet.');
	const dest: CellRange = { start: at, end: { row: at.row + height - 1, col: at.col + width - 1 } };
	const cut = typeof payload !== 'string' && payload.cut && cells.source ? cells.source : undefined;
	// A move rewrites references anywhere in the workbook, so it records the whole workbook.
	const scopes: EditScope[] = cut ? [{ kind: 'workbook' }] : [{ kind: 'sheet', sheet: s }];
	const move = cut && {
		fromSheet: sheetAt(workbook, cut.sheet).name,
		range: cut.range,
		toSheet: sheet.name,
		dRow: at.row - cut.range.start.row,
		dCol: at.col - cut.range.start.col,
	};
	ctx.run(
		cut ? 'Move' : 'Paste',
		'cells',
		scopes,
		() => {
			if (cut && move) {
				rewriteFormulas(workbook, (f, fs) => moveReferencesInFormula(f, fs, move));
				const from = sheetAt(workbook, cut.sheet);
				const doomed: [number, number][] = [];
				forEachCellInRange(from, cut.range, (_c, row, col) => doomed.push([row, col]));
				for (const [row, col] of doomed) deleteCell(from, row, col);
				from.merges = from.merges.filter((m) => !rangeWithin(m, cut.range));
			}
			const styleIds = new Map<object, number>();
			const styleOf = (clip: ClipboardCell): number | undefined => {
				if (!clip.style) return undefined;
				let id = styleIds.get(clip.style);
				if (id === undefined) {
					id = internStyle(workbook, structuredClone(clip.style));
					styleIds.set(clip.style, id);
				}
				return id;
			};
			for (let r = 0; r < cells.rows; r++)
				for (let c = 0; c < cells.cols; c++) {
					const copied = cells.data[r]?.[c] ?? null;
					// A spilled result is recreated by its pasted anchor; only `values` writes it.
					const clip =
						copied?.spilled && mode !== 'values' && mode !== 'formats'
							? { ...copied, value: null, text: '' }
							: copied;
					const row = at.row + (transpose ? c : r);
					const col = at.col + (transpose ? r : c);
					const origin = cells.source && {
						row: cells.source.range.start.row + r,
						col: cells.source.range.start.col + c,
					};
					const formula =
						clip?.formula === undefined
							? undefined
							: move
								? moveReferencesInFormula(clip.formula, move.fromSheet, move, move.toSheet)
								: !origin
									? clip.formula
									: moveFormula(clip.formula, row - origin.row, col - origin.col);
					writeClip(workbook, sheet, row, col, clip, mode, formula, styleOf);
				}
			if (mode === 'all' || mode === 'transpose' || mode === 'formats') {
				sheet.merges = sheet.merges.filter((m) => !rangesIntersect(m, dest));
				for (const m of cells.merges)
					sheet.merges.push(
						transpose
							? {
									start: { row: at.row + m.start.col, col: at.col + m.start.row },
									end: { row: at.row + m.end.col, col: at.col + m.end.row },
								}
							: {
									start: { row: at.row + m.start.row, col: at.col + m.start.col },
									end: { row: at.row + m.end.row, col: at.col + m.end.col },
								},
					);
			}
			if (cut && typeof payload !== 'string') payload.cut = false;
		},
		{ sheet: s, ranges: [dest], ...(cut ? { structural: true } : {}) },
	);
	return dest;
}

function writeClip(
	workbook: Workbook,
	sheet: ReturnType<typeof sheetAt>,
	row: number,
	col: number,
	clip: ClipboardCell | null,
	mode: PasteMode,
	formula: string | undefined,
	styleOf: (clip: ClipboardCell) => number | undefined,
): void {
	if (mode === 'formats') {
		const id = clip ? styleOf(clip) : 0;
		if (id === undefined) return;
		const cell = ensureCell(sheet, row, col);
		if (id) cell.styleId = id;
		else delete cell.styleId;
		pruneCell(sheet, row, col);
		return;
	}
	const existing = getCell(sheet, row, col);
	if (!clip) {
		if (mode === 'all' || mode === 'transpose') deleteCell(sheet, row, col);
		else if (existing) {
			clearContents(existing);
			pruneCell(sheet, row, col);
		}
		return;
	}
	const keepStyle = mode === 'values' || mode === 'formulas' || !clip.style;
	const styleId = keepStyle ? existing?.styleId : styleOf(clip);
	const cell: Cell = { value: clip.value };
	if (styleId) cell.styleId = styleId;
	if (mode !== 'values' && formula !== undefined) {
		cell.formula = formula;
		if (clip.legacyFormula) cell.legacyFormula = true;
	}
	if (clip.richText && mode !== 'values' && mode !== 'formulas')
		cell.richText = structuredClone(clip.richText);
	if (clip.numFmt && styleAt(workbook, cell.styleId).numFmt === 'General') {
		const base = styleAt(workbook, cell.styleId);
		const id = internStyle(workbook, { ...base, numFmt: clip.numFmt });
		if (id) cell.styleId = id;
	}
	putCell(sheet, row, col, cell);
	pruneCell(sheet, row, col);
}
