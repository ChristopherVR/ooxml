import type { CellValue, ConditionalOperator } from '../model.js';
import { isCellError } from '../model.js';

/** Excel's comparison order: numbers < text < booleans. Blank compares as 0 or "". */
function typeRank(value: CellValue): number {
	if (typeof value === 'number') return 0;
	if (typeof value === 'string') return 1;
	if (typeof value === 'boolean') return 2;
	return 3;
}

/** Compares two values the way Excel's `=`/`<` operators do (text case-insensitive). */
export function compareValues(a: CellValue, b: CellValue): number {
	let x = a;
	let y = b;
	if (x === null) x = typeof y === 'string' ? '' : typeof y === 'boolean' ? false : 0;
	if (y === null) y = typeof x === 'string' ? '' : typeof x === 'boolean' ? false : 0;
	const rx = typeRank(x);
	const ry = typeRank(y);
	if (rx !== ry) return rx - ry;
	if (typeof x === 'number' && typeof y === 'number') return x === y ? 0 : x < y ? -1 : 1;
	if (typeof x === 'string' && typeof y === 'string') {
		const lx = x.toLowerCase();
		const ly = y.toLowerCase();
		return lx === ly ? 0 : lx < ly ? -1 : 1;
	}
	if (typeof x === 'boolean' && typeof y === 'boolean') return Number(x) - Number(y);
	return 0;
}

/** Whether `value` satisfies a `cellIs` operator against the evaluated operands. */
export function testOperator(
	operator: ConditionalOperator,
	value: CellValue,
	a: CellValue,
	b: CellValue,
): boolean {
	if (isCellError(value) || isCellError(a)) return false;
	switch (operator) {
		case 'equal':
			return compareValues(value, a) === 0;
		case 'notEqual':
			return compareValues(value, a) !== 0;
		case 'greaterThan':
			return compareValues(value, a) > 0;
		case 'greaterThanOrEqual':
			return compareValues(value, a) >= 0;
		case 'lessThan':
			return compareValues(value, a) < 0;
		case 'lessThanOrEqual':
			return compareValues(value, a) <= 0;
		case 'between':
		case 'notBetween': {
			if (isCellError(b)) return false;
			const lo = compareValues(a, b) <= 0 ? a : b;
			const hi = lo === a ? b : a;
			const inside = compareValues(value, lo) >= 0 && compareValues(value, hi) <= 0;
			return operator === 'between' ? inside : !inside;
		}
	}
}

/** Truthiness of a formula result (`TRUE`, non-zero numbers; text and errors are false). */
export function isTruthy(value: CellValue): boolean {
	if (typeof value === 'boolean') return value;
	if (typeof value === 'number') return value !== 0;
	return false;
}

/** A value reduced to a key for duplicate detection (text case-insensitive). */
export function valueKey(value: CellValue): string | undefined {
	if (value === null || value === '') return undefined;
	if (typeof value === 'number') return `n:${value}`;
	if (typeof value === 'string') return `s:${value.toLowerCase()}`;
	if (typeof value === 'boolean') return `b:${value}`;
	return `e:${value.error}`;
}

/** `PERCENTILE.INC` over ascending `sorted` values; `p` is 0..1. */
export function percentile(sorted: readonly number[], p: number): number {
	if (sorted.length === 0) return 0;
	const rank = Math.max(0, Math.min(1, p)) * (sorted.length - 1);
	const lo = Math.floor(rank);
	const hi = Math.ceil(rank);
	const a = sorted[lo] ?? 0;
	const b = sorted[hi] ?? a;
	return a + (b - a) * (rank - lo);
}

/** Statistics over the numeric cells of a conditional format's ranges. */
export interface RangeStats {
	/** Numeric values ascending. */
	sorted: number[];
	min: number;
	max: number;
	average: number;
	/** Occurrences per {@link valueKey}. */
	counts: Map<string, number>;
}

export function computeStats(values: Iterable<CellValue>): RangeStats {
	const numbers: number[] = [];
	const counts = new Map<string, number>();
	let sum = 0;
	for (const value of values) {
		const key = valueKey(value);
		if (key) counts.set(key, (counts.get(key) ?? 0) + 1);
		if (typeof value === 'number' && Number.isFinite(value)) {
			numbers.push(value);
			sum += value;
		}
	}
	numbers.sort((a, b) => a - b);
	return {
		sorted: numbers,
		min: numbers[0] ?? 0,
		max: numbers.at(-1) ?? 0,
		average: numbers.length ? sum / numbers.length : 0,
		counts,
	};
}

/** Text of a value as Excel's text rules see it. */
export function valueText(value: CellValue): string {
	if (value === null) return '';
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	if (isCellError(value)) return value.error;
	return String(value);
}
