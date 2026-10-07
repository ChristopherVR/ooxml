import {
	type CellAddress,
	type CellRange,
	MAX_COL,
	MAX_ROW,
	normalizeRange,
	rangeContains,
	rangesIntersect,
} from '../address.js';
import { deleteCell, forEachCellInRange, usedRange } from '../cells.js';
import type { Workbook } from '../model.js';
import { internStyle, styleAt } from '../styles.js';
import { toHtml } from './clipboard-html.js';
import { parseHtmlTable } from './clipboard-html-parse.js';
import { parseTsv, toTsv } from './clipboard-text.js';
import { type EditContext, displayText, sheetAt } from './context.js';
import { isSpilledCell, moveReferencesInFormula, parseCellInput } from './deps.js';
import { rewriteFormulas } from './shift-formulas.js';
import { moveFormula } from './fill.js';
import type { EditScope } from './history.js';
import { rangeWithin } from './range-math.js';
import type { ClipboardCell, ClipboardCells, ClipboardPayload, PasteRequest } from './types.js';
import { resolvePasteOptions } from './paste-options.js';
import { writeClip } from './paste-cell.js';
import { copyColumnWidths } from './columns.js';
import { pasteWidths } from './paste-widths.js';
import { copyAnnotations, clearAnnotations, pasteAnnotations } from './clipboard-annotations.js';

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
		row:
			r.start.row === 0 && r.end.row === MAX_ROW
				? Math.max(r.start.row, Math.min(r.end.row, lastRow))
				: r.end.row,
		col:
			r.start.col === 0 && r.end.col === MAX_COL
				? Math.max(r.start.col, Math.min(r.end.col, lastCol))
				: r.end.col,
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
	Object.assign(cells, copyAnnotations(sheet, clipped));
	cells.columnWidthRows = r.end.row - r.start.row + 1;
	cells.columnWidths = copyColumnWidths(sheet, r.start.col, r.end.col);
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
 * `widths` column dimensions only;
 * `transpose` everything with rows and columns swapped. A cut payload clears its source on the
 * first paste and keeps formulas unchanged. Returns the pasted range.
 */
export function pasteAt(
	ctx: EditContext,
	s: number,
	at: CellAddress,
	payload: ClipboardPayload | string,
	request: PasteRequest,
): CellRange {
	const { workbook } = ctx;
	const sheet = sheetAt(workbook, s);
	const cells = typeof payload === 'string' ? cellsFromText(workbook, payload) : payload.cells;
	const { mode, transpose, skipBlanks, operation } = resolvePasteOptions(request);
	if ((mode === 'comments' && !cells.comments) || (mode === 'validation' && !cells.dataValidations))
		throw new RangeError('The clipboard does not contain annotation metadata.');
	if (mode === 'widths')
		return pasteWidths(
			ctx,
			s,
			{ start: at, end: at },
			cells,
			transpose,
			typeof payload !== 'string' && !!payload.cut,
		);
	const height = transpose ? cells.cols : cells.rows;
	const width = transpose ? cells.rows : cells.cols;
	if (!height || !width) return { start: at, end: at };
	if (at.row + height - 1 > MAX_ROW || at.col + width - 1 > MAX_COL)
		throw new RangeError('The paste area extends beyond the sheet.');
	const dest: CellRange = { start: at, end: { row: at.row + height - 1, col: at.col + width - 1 } };
	const cut = typeof payload !== 'string' && payload.cut && cells.source ? cells.source : undefined;
	if (cut && skipBlanks) throw new RangeError('Skip blanks is not available for cut cells.');
	if (cut && mode === 'noBorders')
		throw new RangeError('All except borders is not available for cut cells.');
	if (cut && (mode === 'comments' || mode === 'validation'))
		throw new RangeError('Annotation-only paste is not available for cut cells.');
	if (cut && operation !== 'none')
		throw new RangeError('Paste operations are not available for cut cells.');
	// A move rewrites references anywhere in the workbook: it records the references that change
	// (and every sheet's merges) besides the cells it empties and fills.
	const destCells: EditScope = { kind: 'cells', sheet: s, ranges: [dest] };
	const scopes: EditScope[] = cut
		? [{ kind: 'refs' }, { kind: 'cells', sheet: cut.sheet, ranges: [cut.range] }, destCells]
		: [destCells, { kind: 'parts', sheet: s, parts: ['merges', 'comments', 'dataValidations'] }];
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
				clearAnnotations(from, cut.range);
			}
			pasteAnnotations(
				sheet,
				dest,
				cells,
				mode,
				transpose,
				skipBlanks,
				move ? (f) => moveReferencesInFormula(f, move.fromSheet, move, move.toSheet) : undefined,
			);
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
					// A formula returning empty text is content; a formatted empty cell is blank.
					if (skipBlanks && (!copied || (copied.value === null && copied.formula === undefined)))
						continue;
					// A spilled result is recreated by its pasted anchor; only `values` writes it.
					const clip =
						copied?.spilled && mode !== 'values' && mode !== 'formats'
							? { ...copied, value: null, text: '' }
							: copied;
					const row = at.row + (transpose ? c : r);
					const col = at.col + (transpose ? r : c);
					const origin = cells.source?.range.start;
					const formula =
						clip?.formula === undefined
							? undefined
							: move
								? moveReferencesInFormula(clip.formula, move.fromSheet, move, move.toSheet)
								: !origin
									? clip.formula
									: moveFormula(clip.formula, at.row - origin.row, at.col - origin.col);
					writeClip(workbook, sheet, row, col, clip, mode, formula, styleOf, operation);
				}
			if (mode === 'all' || mode === 'noBorders' || mode === 'formats') {
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
