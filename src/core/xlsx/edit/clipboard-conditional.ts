import { rangesIntersect, type CellRange, type CellAddress } from '../address.js';
import type { ConditionalFormat, Worksheet } from '../model.js';
import type { ClipboardCells } from './types.js';
import { rewriteConditionalRule } from './conditional-formulas.js';
import { rulesByPriority, renumber } from './conditional-formats.js';
import { moveValidationFormula } from './validation-ranges.js';
import { subtractRange } from './range-math.js';

/** Rebase relative formulas when clipping or splitting changes a rule's first range. */
export function rebaseConditionalFormat(
	format: ConditionalFormat,
	ranges: CellRange[],
): ConditionalFormat {
	const next = { ...structuredClone(format), ranges };
	const from = format.ranges[0]?.start;
	const to = ranges[0]?.start;
	if (from && to)
		for (const rule of next.rules)
			rewriteConditionalRule(rule, (f) =>
				moveValidationFormula(f, to.row - from.row, to.col - from.col),
			);
	return next;
}

export function removeConditionalArea(sheet: Worksheet, area: CellRange): void {
	sheet.conditionalFormats = sheet.conditionalFormats.flatMap((format) => {
		const ranges = format.ranges.flatMap((r) => subtractRange(r, area));
		return ranges.length ? [rebaseConditionalFormat(format, ranges)] : [];
	});
}

export function copyConditionalFormats(sheet: Worksheet, area: CellRange): ConditionalFormat[] {
	const point = (at: CellAddress) => ({
		row: at.row - area.start.row,
		col: at.col - area.start.col,
	});
	return sheet.conditionalFormats.flatMap((format) => {
		const ranges = format.ranges
			.filter((r) => rangesIntersect(r, area))
			.map((r) => ({
				start: {
					row: Math.max(r.start.row, area.start.row),
					col: Math.max(r.start.col, area.start.col),
				},
				end: { row: Math.min(r.end.row, area.end.row), col: Math.min(r.end.col, area.end.col) },
			}));
		if (!ranges.length) return [];
		return [
			{
				...rebaseConditionalFormat(format, ranges),
				ranges: ranges.map((r) => ({ start: point(r.start), end: point(r.end) })),
			},
		];
	});
}

/** Skip Blanks replaces rules only where copied rules apply, independently of cell values. */
export function pasteConditionalFormats(
	sheet: Worksheet,
	dest: CellRange,
	cells: ClipboardCells,
	transpose: boolean,
	skipBlanks: boolean,
	cutFormula?: (formula: string) => string,
	merge = false,
): void {
	if (!cells.conditionalFormats) return;
	const point = (at: CellAddress) => ({
		row: dest.start.row + (transpose ? at.col : at.row),
		col: dest.start.col + (transpose ? at.row : at.col),
	});
	const copied = cells.conditionalFormats.map((format) => {
		const next = structuredClone(format);
		// Preserved x14 records belong to the original rule; a new base rule must not share its ID.
		for (const rule of next.rules) if (rule.type === 'dataBar') delete rule.extensionId;
		const height = transpose ? cells.cols : cells.rows;
		const width = transpose ? cells.rows : cells.cols;
		next.ranges = format.ranges.flatMap((r) => {
			const mapped = { start: point(r.start), end: point(r.end) };
			const fullRows =
				mapped.start.row === dest.start.row && mapped.end.row === dest.start.row + height - 1;
			const fullCols =
				mapped.start.col === dest.start.col && mapped.end.col === dest.start.col + width - 1;
			const ranges: CellRange[] = [];
			for (
				let col = dest.start.col;
				col <= dest.end.col;
				col += fullCols ? dest.end.col - dest.start.col + 1 : width
			)
				for (
					let row = dest.start.row;
					row <= dest.end.row;
					row += fullRows ? dest.end.row - dest.start.row + 1 : height
				)
					ranges.push({
						start: {
							row: mapped.start.row + row - dest.start.row,
							col: mapped.start.col + col - dest.start.col,
						},
						end: {
							row: fullRows ? dest.end.row : mapped.end.row + row - dest.start.row,
							col: fullCols ? dest.end.col : mapped.end.col + col - dest.start.col,
						},
					});
			return ranges;
		});
		const source = format.ranges[0]?.start;
		const target = next.ranges[0]?.start;
		const origin = cells.source?.range.start ?? { row: 0, col: 0 };
		if (source && target)
			for (const rule of next.rules)
				rewriteConditionalRule(
					rule,
					cutFormula ??
						((f) =>
							moveValidationFormula(
								f,
								target.row - origin.row - source.row,
								target.col - origin.col - source.col,
							)),
				);
		return next;
	});
	if (!skipBlanks && !merge) removeConditionalArea(sheet, dest);
	else if (skipBlanks)
		for (const format of copied)
			for (const range of format.ranges) removeConditionalArea(sheet, range);
	const original = rulesByPriority(sheet);
	const added = rulesByPriority({ ...sheet, conditionalFormats: copied });
	sheet.conditionalFormats.push(...copied);
	renumber([...added, ...original]);
}
