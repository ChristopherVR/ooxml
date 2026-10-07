import type { CellAddress, CellRange } from '../address.js';
import type { DataValidation } from '../model.js';
import { moveFormula } from './fill.js';
import { subtractRange } from './range-math.js';

/** Validation formulas are relative to the top-left of their ranges' bounding box. */
export function validationAnchor(ranges: CellRange[]): CellAddress {
	let row = Infinity;
	let col = Infinity;
	for (const r of ranges) {
		row = Math.min(row, r.start.row, r.end.row);
		col = Math.min(col, r.start.col, r.end.col);
	}
	return { row, col };
}

export function moveValidationFormula(formula: string, dRow: number, dCol: number): string {
	const equals = formula.startsWith('=');
	const moved = moveFormula(equals ? formula.slice(1) : formula, dRow, dCol);
	return equals ? `=${moved}` : moved;
}

/** Replaces ranges while retaining each surviving cell's relative validation semantics. */
export function rebaseValidation(rule: DataValidation, ranges: CellRange[]): DataValidation {
	const old = validationAnchor(rule.ranges);
	const next = validationAnchor(ranges);
	const result = { ...structuredClone(rule), ranges };
	for (const key of ['formula1', 'formula2'] as const)
		if (result[key] !== undefined)
			result[key] = moveValidationFormula(result[key], next.row - old.row, next.col - old.col);
	return result;
}

/** Removes an area from validation rules, adjusting formulas when their origin changes. */
export function removeValidationArea(rules: DataValidation[], area: CellRange): DataValidation[] {
	return rules.flatMap((rule) => {
		const ranges = rule.ranges.flatMap((r) => subtractRange(r, area));
		return ranges.length ? [rebaseValidation(rule, ranges)] : [];
	});
}
