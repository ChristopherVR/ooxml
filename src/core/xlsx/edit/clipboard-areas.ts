// Copying a multi-area selection. Excel (checked with Excel 16 through COM) allows it only when
// every area has the same rows (side by side) or the same columns (stacked), and then puts the
// areas next to each other on the clipboard with the gaps between them dropped, in sheet order;
// any other shape is refused with "This action won't work on multiple selections."
// Validations and conditional formats of the areas travel with their ranges moved the same way.
// Excel refuses to cut several areas (see CUT_MULTIPLE).
import { type CellRange, normalizeRange, rangesIntersect } from '../address';
import type { Workbook } from '../model';
import { cellsFromText, copyRange } from './clipboard';
import { toHtml } from './clipboard-html';
import { toTsv } from './clipboard-text';
import { rewriteConditionalRule } from './conditional-formulas';
import { moveFormula } from './fill';
import { moveValidationFormula } from './validation-ranges';
import { resolvePasteOptions } from './paste-options';
import type {
	ClipboardCell,
	ClipboardCells,
	ClipboardPayload,
	EditSession,
	PasteRequest,
} from './types';

export const MULTIPLE_SELECTIONS = "This action won't work on multiple selections.";

/** Excel's message for cutting several areas. */
export const CUT_MULTIPLE =
	"This can't be done on a multiple range selection. Select a single range and try again.";

/** How the areas line up: `across` share their rows, `down` share their columns. */
export function areaLayout(ranges: readonly CellRange[]): 'across' | 'down' | undefined {
	const [first, ...rest] = ranges.map(normalizeRange);
	if (!first) return undefined;
	for (let i = 0; i < rest.length; i++)
		for (let j = 0; j < i; j++) {
			const a = rest[i];
			const b = rest[j];
			if (a && b && rangesIntersect(a, b)) return undefined;
		}
	if (rest.some((r) => rangesIntersect(first, r))) return undefined;
	if (rest.every((r) => r.start.row === first.start.row && r.end.row === first.end.row))
		return 'across';
	if (rest.every((r) => r.start.col === first.start.col && r.end.col === first.end.col))
		return 'down';
	return undefined;
}

/** Copies several areas as one block; throws a RangeError for a shape Excel refuses. */
export function copyAreas(
	workbook: Workbook,
	s: number,
	ranges: readonly CellRange[],
): ClipboardPayload {
	const only = ranges[0];
	if (ranges.length === 1 && only) return copyRange(workbook, s, only);
	const layout = areaLayout(ranges);
	if (!layout) throw new RangeError(MULTIPLE_SELECTIONS);
	const across = layout === 'across';
	const areas = ranges
		.map(normalizeRange)
		.sort((a, b) => (across ? a.start.col - b.start.col : a.start.row - b.start.row));
	const parts = areas.map((area) => ({ area, copy: copyRange(workbook, s, area) }));
	const first = parts[0];
	if (!first) throw new RangeError(MULTIPLE_SELECTIONS);
	const rows = across ? first.copy.cells.rows : parts.reduce((n, p) => n + p.copy.cells.rows, 0);
	const cols = across ? parts.reduce((n, p) => n + p.copy.cells.cols, 0) : first.copy.cells.cols;
	const data: (ClipboardCell | null)[][] = Array.from({ length: rows }, () =>
		Array.from({ length: cols }, () => null),
	);
	const cells: ClipboardCells = { rows, cols, data, merges: [] };
	const comments: NonNullable<ClipboardCells['comments']> = [];
	const hyperlinks: NonNullable<ClipboardCells['hyperlinks']> = [];
	const validations: NonNullable<ClipboardCells['dataValidations']> = [];
	const formats: NonNullable<ClipboardCells['conditionalFormats']> = [];
	const widths: number[] = [];
	let offset = 0;
	for (const { area, copy } of parts) {
		const dRow = across ? 0 : offset;
		const dCol = across ? offset : 0;
		// A paste moves formulas by (target - first area); an area that moved by a different
		// amount (the dropped gap) is corrected here so every cell keeps its own relative refs.
		const gap = across
			? area.start.col - first.area.start.col - offset
			: area.start.row - first.area.start.row - offset;
		copy.cells.data.forEach((line, r) =>
			line.forEach((cell, c) => {
				const target = data[r + dRow];
				if (!target) return;
				target[c + dCol] =
					cell?.formula === undefined || gap === 0
						? cell
						: {
								...cell,
								formula: across
									? moveFormula(cell.formula, 0, -gap)
									: moveFormula(cell.formula, -gap, 0),
							};
			}),
		);
		const at = (a: { row: number; col: number }) => ({ row: a.row + dRow, col: a.col + dCol });
		for (const m of copy.cells.merges) cells.merges.push({ start: at(m.start), end: at(m.end) });
		for (const c of copy.cells.comments ?? [])
			comments.push({ ...structuredClone(c), address: at(c.address) });
		for (const h of copy.cells.hyperlinks ?? [])
			hyperlinks.push({
				...structuredClone(h),
				range: { start: at(h.range.start), end: at(h.range.end) },
			});
		const moveRange = (r: CellRange): CellRange => ({ start: at(r.start), end: at(r.end) });
		const fix = (f: string): string =>
			gap === 0
				? f
				: across
					? moveValidationFormula(f, 0, -gap)
					: moveValidationFormula(f, -gap, 0);
		for (const v of copy.cells.dataValidations ?? []) {
			const next = { ...structuredClone(v), ranges: v.ranges.map(moveRange) };
			if (next.formula1 !== undefined) next.formula1 = fix(next.formula1);
			if (next.formula2 !== undefined) next.formula2 = fix(next.formula2);
			validations.push(next);
		}
		for (const f of copy.cells.conditionalFormats ?? []) {
			const next = { ...structuredClone(f), ranges: f.ranges.map(moveRange) };
			if (gap !== 0) for (const rule of next.rules) rewriteConditionalRule(rule, fix);
			formats.push(next);
		}
		if (across) widths.push(...(copy.cells.columnWidths ?? []));
		offset += across ? copy.cells.cols : copy.cells.rows;
	}
	if (comments.length) cells.comments = comments;
	if (hyperlinks.length) cells.hyperlinks = hyperlinks;
	if (validations.length) cells.dataValidations = validations;
	if (formats.length) cells.conditionalFormats = formats;
	cells.columnWidths = across ? widths : (first.copy.cells.columnWidths ?? []);
	cells.columnWidthRows = rows;
	cells.source = {
		sheet: s,
		range: {
			start: first.area.start,
			end: { row: first.area.start.row + rows - 1, col: first.area.start.col + cols - 1 },
		},
	};
	return {
		tsv: toTsv(data.map((line) => line.map((c) => c?.text ?? ''))),
		html: toHtml(cells, workbook.theme),
		cells,
	};
}

/**
 * Pastes a payload into every area of a multi-area selection as one undo step (Excel 16 pastes
 * the block into each area: a single cell receives it at its top-left, a larger area must hold
 * whole copies). A cut cannot go into several areas; nothing is written if any area does not fit.
 */
export function pasteAreas(
	session: EditSession,
	sheet: number,
	ranges: readonly CellRange[],
	payload: ClipboardPayload | string,
	mode: PasteRequest = 'all',
): CellRange {
	const areas = ranges.map(normalizeRange);
	if (typeof payload !== 'string' && payload.cut) throw new RangeError(CUT_MULTIPLE);
	const clip =
		typeof payload === 'string' ? cellsFromText(session.workbook, payload) : payload.cells;
	const turned = resolvePasteOptions(mode).transpose;
	const height = turned ? clip.cols : clip.rows;
	const width = turned ? clip.rows : clip.cols;
	for (const area of areas) {
		const rows = area.end.row - area.start.row + 1;
		const cols = area.end.col - area.start.col + 1;
		if ((rows > 1 || cols > 1) && height > 0 && width > 0 && (rows % height || cols % width))
			throw new RangeError('The copy and paste areas are not compatible sizes.');
	}
	let last = areas[0];
	session.batch('Paste', () => {
		for (const area of areas) last = session.paste(sheet, area, payload, mode);
	});
	return last ?? { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };
}
