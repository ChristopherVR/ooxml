import type { CellRange } from '../address.js';
import type { EvalHost } from './context.js';
import { toMatrix } from './references.js';
import {
	type Area,
	ERR,
	LambdaValue,
	Matrix,
	RefValue,
	type Scalar,
	type Value,
} from './values.js';

/** The value a formula cell shows: one value, or an array to spill. Blanks become 0. */
export function finalize(value: Value, host: EvalHost): Scalar | Matrix {
	let result: Scalar | Matrix;
	if (value instanceof RefValue) {
		if (value.areas.length !== 1) return ERR.VALUE;
		if (value.isCell()) {
			const area = value.areas[0] as Area;
			result = host.readCell(area.sheet, area.range.start.row, area.range.start.col);
		} else result = toMatrix(host, value);
	} else if (value instanceof LambdaValue) return ERR.CALC;
	else result = value;
	if (result instanceof Matrix) {
		if (result.rows === 1 && result.cols === 1) result = result.get(0, 0);
		else return result.map(clean);
	}
	return clean(result);
}

export function clean(value: Scalar): Scalar {
	if (value === null) return 0;
	if (typeof value === 'number') {
		if (!Number.isFinite(value)) return ERR.NUM;
		return value === 0 ? 0 : value;
	}
	return value;
}

export const rangeHas = (range: CellRange, row: number, col: number): boolean =>
	row >= range.start.row && row <= range.end.row && col >= range.start.col && col <= range.end.col;

export function sameRange(a: CellRange | undefined, b: CellRange | undefined): boolean {
	if (!a || !b) return a === b || (!a && !b);
	return (
		a.start.row === b.start.row &&
		a.start.col === b.start.col &&
		a.end.row === b.end.row &&
		a.end.col === b.end.col
	);
}

export function boundingBox(ranges: CellRange[]): CellRange {
	return {
		start: {
			row: Math.min(...ranges.map((r) => r.start.row)),
			col: Math.min(...ranges.map((r) => r.start.col)),
		},
		end: {
			row: Math.max(...ranges.map((r) => r.end.row)),
			col: Math.max(...ranges.map((r) => r.end.col)),
		},
	};
}
