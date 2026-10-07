import { rangeContains, rangesIntersect, type CellAddress, type CellRange } from '../address.js';
import type { Worksheet } from '../model.js';
import {
	removeValidationArea,
	rebaseValidation,
	validationAnchor,
	moveValidationFormula,
} from './validation-ranges.js';
import type { ClipboardCells, PasteMode } from './types.js';

export function copyAnnotations(
	sheet: Worksheet,
	area: CellRange,
): Pick<ClipboardCells, 'comments' | 'dataValidations'> {
	const relative = (at: CellAddress) => ({
		row: at.row - area.start.row,
		col: at.col - area.start.col,
	});
	return {
		comments: sheet.comments
			.filter((c) => rangeContains(area, c.address))
			.map((c) => ({ ...structuredClone(c), address: relative(c.address) })),
		dataValidations: sheet.dataValidations.flatMap((rule) => {
			const ranges = rule.ranges
				.filter((r) => rangesIntersect(r, area))
				.map((r) => ({
					start: {
						row: Math.max(r.start.row, area.start.row),
						col: Math.max(r.start.col, area.start.col),
					},
					end: { row: Math.min(r.end.row, area.end.row), col: Math.min(r.end.col, area.end.col) },
				}));
			return ranges.length
				? [
						{
							...rebaseValidation(rule, ranges),
							ranges: ranges.map((r) => ({ start: relative(r.start), end: relative(r.end) })),
						},
					]
				: [];
		}),
	};
}

export function clearAnnotations(sheet: Worksheet, area: CellRange): void {
	sheet.comments = sheet.comments.filter((c) => !rangeContains(area, c.address));
	sheet.dataValidations = removeValidationArea(sheet.dataValidations, area);
}

/** Metadata blanks are independent of cell-value blanks in Excel's Skip Blanks operation. */
export function pasteAnnotations(
	sheet: Worksheet,
	dest: CellRange,
	cells: ClipboardCells,
	mode: PasteMode,
	transpose: boolean,
	skipBlanks: boolean,
	cutFormula?: (formula: string) => string,
): void {
	const point = (at: CellAddress) => ({
		row: dest.start.row + (transpose ? at.col : at.row),
		col: dest.start.col + (transpose ? at.row : at.col),
	});
	if (
		(mode === 'all' || mode === 'mergeFormats' || mode === 'noBorders' || mode === 'comments') &&
		cells.comments
	) {
		if (!skipBlanks) sheet.comments = sheet.comments.filter((c) => !rangeContains(dest, c.address));
		const key = (at: CellAddress) => `${at.row}:${at.col}`;
		const comments = new Map(sheet.comments.map((c) => [key(c.address), c]));
		for (const comment of cells.comments) {
			const address = point(comment.address);
			comments.delete(key(address));
			comments.set(key(address), { ...structuredClone(comment), address });
		}
		sheet.comments = [...comments.values()];
	}
	if (
		(mode === 'all' || mode === 'mergeFormats' || mode === 'noBorders' || mode === 'validation') &&
		cells.dataValidations
	) {
		if (!skipBlanks) sheet.dataValidations = removeValidationArea(sheet.dataValidations, dest);
		for (const rule of cells.dataValidations) {
			const ranges = rule.ranges.map((r) => ({ start: point(r.start), end: point(r.end) }));
			if (skipBlanks)
				for (const r of ranges)
					sheet.dataValidations = removeValidationArea(sheet.dataValidations, r);
			const source = validationAnchor(rule.ranges);
			const target = validationAnchor(ranges);
			const origin = cells.source?.range.start ?? { row: 0, col: 0 };
			const pasted = { ...structuredClone(rule), ranges };
			for (const key of ['formula1', 'formula2'] as const)
				if (pasted[key] !== undefined)
					pasted[key] = cutFormula
						? cutFormula(pasted[key])
						: moveValidationFormula(
								pasted[key],
								target.row - origin.row - source.row,
								target.col - origin.col - source.col,
							);
			sheet.dataValidations.push(pasted);
		}
	}
}
