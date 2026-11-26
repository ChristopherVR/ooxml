import {
	type CellAddress,
	type CellRange,
	MAX_COL,
	MAX_ROW,
	normalizeRange,
	rangeContains,
	rangesIntersect,
} from '../address';
import { deleteCell, forEachCellInRange, usedRange } from '../cells';
import type { CellStyle, Workbook } from '../model';
import { internStyle, styleAt } from '../styles';
import { toHtml } from './clipboard-html';
import { parseHtmlTable } from './clipboard-html-parse';
import { parseTsv, toTsv } from './clipboard-text';
import { type EditContext, displayText, sheetAt } from './context';
import { isSpilledCell, moveReferencesInFormula, parseCellInput } from './deps';
import { rewriteFormulas } from './shift-formulas';
import { moveFormula } from './fill';
import type { EditScope } from './history';
import { rangeWithin } from './range-math';
import type { ClipboardCell, ClipboardCells, ClipboardPayload, PasteRequest } from './types';
import { resolvePasteOptions } from './paste-options';
import { writeClip } from './paste-cell';
import { syncTableHeader, writeScopes } from './cell-values';
import { copyColumnWidths } from './columns';
import { pasteWidths } from './paste-widths';
import { copyAnnotations, clearAnnotations, pasteAnnotations } from './clipboard-annotations';
import { copyHyperlinks, clearHyperlinks, pasteHyperlinks } from './clipboard-links';
import {
	copyConditionalFormats,
	removeConditionalArea,
	pasteConditionalFormats,
} from './clipboard-conditional';

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
	// Cells with the same format share one copy of it: a payload's styles are read, never changed,
	// so a large copy clones each format once, and a paste (which keys its interning on the style
	// object) interns each once.
	const styles = new Map<number, CellStyle>();
	const styleCopy = (id: number | undefined): CellStyle => {
		let style = styles.get(id ?? 0);
		if (!style) styles.set(id ?? 0, (style = structuredClone(styleAt(workbook, id))));
		return style;
	};
	forEachCellInRange(sheet, clipped, (cell, row, col) => {
		const clip: ClipboardCell = {
			value: cell.value,
			style: styleCopy(cell.styleId),
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
	cells.hyperlinks = copyHyperlinks(sheet, clipped);
	cells.conditionalFormats = copyConditionalFormats(sheet, clipped);
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
 * first paste and keeps formulas unchanged. A repeat range tiles the source in one undo step.
 */
export function pasteAt(
	ctx: EditContext,
	s: number,
	at: CellAddress,
	payload: ClipboardPayload | string,
	request: PasteRequest,
	repeat?: CellRange,
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
	const end = { row: at.row + height - 1, col: at.col + width - 1 };
	if (!repeat && (end.row > MAX_ROW || end.col > MAX_COL))
		throw new RangeError('The paste area extends beyond the sheet.');
	const dest = repeat ?? { start: at, end };
	const cut = typeof payload !== 'string' && payload.cut && cells.source ? cells.source : undefined;
	if (cut && skipBlanks) throw new RangeError('Skip blanks is not available for cut cells.');
	if (cut && (mode === 'noBorders' || mode === 'mergeFormats'))
		throw new RangeError('This paste mode is not available for cut cells.');
	if (cut && (mode === 'comments' || mode === 'validation'))
		throw new RangeError('Annotation-only paste is not available for cut cells.');
	if (cut && operation !== 'none')
		throw new RangeError('Paste operations are not available for cut cells.');
	// A move rewrites references anywhere in the workbook: it records the references that change
	// (and every sheet's merges) besides the cells it empties and fills.
	const destCells: EditScope = { kind: 'cells', sheet: s, ranges: [dest] };
	const scopes: EditScope[] = cut
		? [{ kind: 'refs' }, { kind: 'cells', sheet: cut.sheet, ranges: [cut.range] }, destCells]
		: writeScopes(sheet, s, dest, [
				'merges',
				'comments',
				'dataValidations',
				'hyperlinks',
				'conditionalFormats',
			]);
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
				const from = sheetAt(workbook, cut.sheet);
				// Rebase the surviving rule before reference rewrites see its old, moved anchor.
				removeConditionalArea(from, cut.range);
				rewriteFormulas(workbook, (f, fs) => moveReferencesInFormula(f, fs, move));
				const doomed: [number, number][] = [];
				forEachCellInRange(from, cut.range, (_c, row, col) => doomed.push([row, col]));
				for (const [row, col] of doomed) deleteCell(from, row, col);
				from.merges = from.merges.filter((m) => !rangeWithin(m, cut.range));
				clearAnnotations(from, cut.range);
				clearHyperlinks(from, cut.range);
			}
			const cutFormula = move
				? (f: string) => moveReferencesInFormula(f, move.fromSheet, move, move.toSheet)
				: undefined;
			const carriesLinks = mode === 'all' || mode === 'mergeFormats' || mode === 'noBorders';
			const carriesFormats = carriesLinks || mode === 'formats';
			const hasAnnotations =
				(cells.comments && (cells.comments.length || sheet.comments.length)) ||
				(cells.dataValidations && (cells.dataValidations.length || sheet.dataValidations.length));
			const hasLinks =
				carriesLinks && cells.hyperlinks && (cells.hyperlinks.length || sheet.hyperlinks.length);
			const styleIds = new Map<CellStyle, number>();
			const styleOf = (clip: ClipboardCell): number | undefined => {
				if (!clip.style) return undefined;
				let id = styleIds.get(clip.style);
				if (id === undefined) {
					id = internStyle(workbook, structuredClone(clip.style));
					styleIds.set(clip.style, id);
				}
				return id;
			};
			const origin = cells.source?.range.start;
			const syncHeaders =
				sheet.tables.length > 0 &&
				mode !== 'formats' &&
				mode !== 'comments' &&
				mode !== 'validation';
			const last = repeat?.end ?? at;
			for (let top = at.row; top <= last.row; top += height)
				for (let left = at.col; left <= last.col; left += width) {
					// A repeated selection is one undo step. Validate each tile so an out-of-bounds
					// tile takes the same rollback path as other failures partway through the paste.
					const bottom = top + height - 1;
					const right = left + width - 1;
					if (bottom > MAX_ROW || right > MAX_COL)
						throw new RangeError('The paste area extends beyond the sheet.');
					const movesMerges = carriesFormats && (sheet.merges.length || cells.merges.length);
					const tile =
						hasAnnotations || hasLinks || movesMerges
							? { start: { row: top, col: left }, end: { row: bottom, col: right } }
							: undefined;
					if (tile && hasAnnotations)
						pasteAnnotations(sheet, tile, cells, mode, transpose, skipBlanks, cutFormula);
					if (tile && hasLinks)
						pasteHyperlinks(sheet, tile, cells, transpose, skipBlanks, cutFormula);
					if (!repeat && carriesFormats)
						pasteConditionalFormats(
							sheet,
							dest,
							cells,
							transpose,
							skipBlanks,
							cutFormula,
							mode === 'mergeFormats',
						);
					for (let r = 0; r < cells.rows; r++)
						for (let c = 0; c < cells.cols; c++) {
							const copied = cells.data[r]?.[c] ?? null;
							// A formula returning empty text is content; a formatted empty cell is blank.
							if (
								skipBlanks &&
								(!copied || (copied.value === null && copied.formula === undefined))
							)
								continue;
							// A spilled result is recreated by its pasted anchor; only values writes it.
							const clip =
								copied?.spilled && mode !== 'values' && mode !== 'formats'
									? { ...copied, value: null, text: '' }
									: copied;
							const row = top + (transpose ? c : r);
							const col = left + (transpose ? r : c);
							const formula =
								clip?.formula === undefined
									? undefined
									: cutFormula
										? cutFormula(clip.formula)
										: origin
											? moveFormula(clip.formula, top - origin.row, left - origin.col)
											: clip.formula;
							writeClip(workbook, sheet, row, col, clip, mode, formula, styleOf, operation);
							if (syncHeaders) syncTableHeader(ctx, sheet, row, col);
						}
					if (tile && movesMerges) {
						sheet.merges = sheet.merges.filter((m) => !rangesIntersect(m, tile));
						const point = (cell: CellAddress): CellAddress => ({
							row: top + (transpose ? cell.col : cell.row),
							col: left + (transpose ? cell.row : cell.col),
						});
						for (const m of cells.merges)
							sheet.merges.push({ start: point(m.start), end: point(m.end) });
					}
				}
			if (repeat && carriesFormats)
				pasteConditionalFormats(
					sheet,
					dest,
					cells,
					transpose,
					skipBlanks,
					undefined,
					mode === 'mergeFormats',
				);
			if (cut && typeof payload !== 'string') payload.cut = false;
		},
		{ sheet: s, ranges: [dest], ...(cut ? { structural: true } : {}) },
	);
	return dest;
}
