import type { CellRange } from '../address.js';
import { type RefSpec, type Token, formatRefSpec, tokenize } from './deps.js';
import { type AxisShift, sameRange, shiftRangeInBand } from './range-math.js';
import type { Band } from './shift-sheet.js';

/** Whether a reference token points at `sheet` (unqualified ones point at the formula's sheet). */
function pointsAt(token: Token, formulaSheet: string, sheet: string): boolean {
	const prefix = token.prefix;
	if (!prefix) return formulaSheet.toLowerCase() === sheet.toLowerCase();
	return (
		prefix.book === undefined &&
		prefix.sheet2 === undefined &&
		prefix.sheet.toLowerCase() === sheet.toLowerCase()
	);
}

/**
 * Rewrites a formula for an insert-cells or delete-cells shift: only cell and area references lying
 * entirely inside the band move (as Excel does); references to deleted cells become `#REF!`.
 * Unparseable formulas are returned unchanged.
 */
export function shiftFormulaInBand(
	formula: string,
	formulaSheet: string,
	sheet: string,
	shift: AxisShift,
	band: Band,
): string {
	let tokens: Token[];
	try {
		tokens = tokenize(formula);
	} catch {
		return formula;
	}
	let changed = false;
	for (const token of tokens) {
		const ref = token.ref;
		if (token.kind !== 'ref' || !ref || (ref.kind !== 'cell' && ref.kind !== 'area')) continue;
		if (!pointsAt(token, formulaSheet, sheet)) continue;
		const rowFirst = ref.start.row <= ref.end.row;
		const colFirst = ref.start.col <= ref.end.col;
		const range: CellRange = {
			start: {
				row: Math.min(ref.start.row, ref.end.row),
				col: Math.min(ref.start.col, ref.end.col),
			},
			end: { row: Math.max(ref.start.row, ref.end.row), col: Math.max(ref.start.col, ref.end.col) },
		};
		const moved = shiftRangeInBand(range, shift, band.lo, band.hi);
		if (moved && sameRange(moved, range)) continue;
		const prefix = token.prefix?.text ?? '';
		if (!moved) {
			token.text = `${prefix}#REF!`;
			changed = true;
			continue;
		}
		const next: RefSpec = {
			kind: ref.kind,
			start: {
				...ref.start,
				row: rowFirst ? moved.start.row : moved.end.row,
				col: colFirst ? moved.start.col : moved.end.col,
			},
			end: {
				...ref.end,
				row: rowFirst ? moved.end.row : moved.start.row,
				col: colFirst ? moved.end.col : moved.start.col,
			},
		};
		if (ref.kind === 'cell') next.end = next.start;
		token.text = `${prefix}${formatRefSpec(next)}${token.spill ? '#' : ''}`;
		changed = true;
	}
	return changed ? tokens.map((t) => t.text).join('') : formula;
}
