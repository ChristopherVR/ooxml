// Excel's scalar coercion rules.
import type { CellError } from '../model.js';
import { numberToText, parseNumberText, round15 } from './text-number.js';
import { ERR, ErrorSignal, isError, type Scalar } from './values.js';

/** A number as arithmetic sees it: blank is 0, TRUE is 1, numeric text converts. Throws on errors. */
export function toNumber(value: Scalar): number {
	if (typeof value === 'number') return value;
	if (value === null) return 0;
	if (typeof value === 'boolean') return value ? 1 : 0;
	if (typeof value === 'string') {
		const parsed = parseNumberText(value);
		if (parsed === undefined) throw new ErrorSignal(ERR.VALUE);
		return parsed;
	}
	throw new ErrorSignal(value);
}

/** Text as `&` sees it: blank is empty, numbers use General formatting. Throws on errors. */
export function toText(value: Scalar): string {
	if (typeof value === 'string') return value;
	if (typeof value === 'number') return numberToText(value);
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	if (value === null) return '';
	throw new ErrorSignal(value);
}

/** A logical as IF sees it: numbers are non-zero, text must be TRUE or FALSE. Throws on errors. */
export function toBool(value: Scalar): boolean {
	if (typeof value === 'boolean') return value;
	if (typeof value === 'number') return value !== 0;
	if (value === null) return false;
	if (typeof value === 'string') {
		const upper = value.toUpperCase();
		if (upper === 'TRUE') return true;
		if (upper === 'FALSE') return false;
		throw new ErrorSignal(ERR.VALUE);
	}
	throw new ErrorSignal(value);
}

/** Throws when `value` is an error. */
export function check(value: Scalar): Exclude<Scalar, CellError> {
	if (isError(value)) throw new ErrorSignal(value);
	return value;
}

const TYPE_RANK = (value: Exclude<Scalar, CellError | null>): number =>
	typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2;

/** Compares numbers at Excel's 15-significant-digit precision. */
export function compareNumbers(a: number, b: number): number {
	if (a === b) return 0;
	// Numbers further apart than one unit in the 15th digit cannot round to the same value.
	if (Math.abs(a - b) > Math.max(Math.abs(a), Math.abs(b)) * 1e-14) return a < b ? -1 : 1;
	const ra = round15(a);
	const rb = round15(b);
	return ra === rb ? 0 : ra < rb ? -1 : 1;
}

/** Case-insensitive text comparison. */
export function compareText(a: string, b: string): number {
	const la = a.toLowerCase();
	const lb = b.toLowerCase();
	return la === lb ? 0 : la < lb ? -1 : 1;
}

/**
 * Compares two non-error scalars like Excel's comparison operators: numbers sort before text,
 * text before logicals; blank takes the other side's type (0, "", FALSE).
 */
export function compareScalars(
	a: Exclude<Scalar, CellError>,
	b: Exclude<Scalar, CellError>,
): number {
	if (a === null && b === null) return 0;
	if (a === null) return compareScalars(blankLike(b), b);
	if (b === null) return compareScalars(a, blankLike(a));
	const ta = TYPE_RANK(a);
	const tb = TYPE_RANK(b);
	if (ta !== tb) return ta < tb ? -1 : 1;
	if (typeof a === 'number' && typeof b === 'number') return compareNumbers(a, b);
	if (typeof a === 'string' && typeof b === 'string') return compareText(a, b);
	return a === b ? 0 : a ? 1 : -1;
}

function blankLike(other: Exclude<Scalar, CellError>): Exclude<Scalar, CellError> {
	if (typeof other === 'number') return 0;
	if (typeof other === 'string') return '';
	if (typeof other === 'boolean') return false;
	return null;
}

/** A finite result, or `#NUM!`. */
export function finite(value: number): number {
	if (!Number.isFinite(value)) throw new ErrorSignal(ERR.NUM);
	return value;
}
