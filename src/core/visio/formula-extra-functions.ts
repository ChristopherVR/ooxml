import { formulaFailure } from './formula';
import type { VisioFormulaValue } from './formula';

/**
 * Remaining ShapeSheet functions with a bounded, documented subset: STRSAMEEX comparison flags,
 * FORMAT pictures, DATE and the angle and point conversions between a shape and its parent.
 * Anything outside the subset is refused as unsupported instead of guessed.
 * https://learn.microsoft.com/en-us/office/client-developer/visio/strsameex-function
 * https://learn.microsoft.com/en-us/office/client-developer/visio/format-function
 */

const IGNORE_CASE = 1,
	IGNORE_NONSPACE = 2,
	IGNORE_SYMBOLS = 4;

/** STRSAMEEX flags are the Windows NORM_* bits; case, nonspacing marks and symbols are handled. */
export function strSameEx(a: string, b: string, flags: number): boolean {
	if (!Number.isInteger(flags) || flags < 0)
		return formulaFailure('value', 'STRSAMEEX flags must be a whole number');
	if (flags & ~(IGNORE_CASE | IGNORE_NONSPACE | IGNORE_SYMBOLS))
		return formulaFailure('unsupported', 'STRSAMEEX with kana or width comparison flags');
	const fold = (text: string) => {
		let result = text;
		if (flags & IGNORE_NONSPACE) result = result.normalize('NFD').replace(/\p{Mn}/gu, '');
		if (flags & IGNORE_SYMBOLS) result = result.replace(/[\p{P}\p{S}\s]/gu, '');
		return flags & IGNORE_CASE ? result.toLowerCase() : result;
	};
	return fold(a) === fold(b);
}

/**
 * FORMAT(number, picture) for numeric pictures made of `0`, `#`, `.`, `,` and a trailing `%`
 * (for example `0.00`, `#,##0.0`, `0%`). Section lists, dates, exponents and literals are refused.
 */
export function formatNumberPicture(value: VisioFormulaValue, picture: string): string {
	if (value.unit !== 'scalar') return formulaFailure('unit', 'FORMAT requires a plain number');
	const match = /^([#,0]*)(?:\.([0#]*))?(%?)$/.exec(picture);
	const integerPicture = match?.[1] ?? '';
	if (
		!match ||
		!integerPicture.includes('0') ||
		/0,*#/.test(integerPicture) ||
		integerPicture.startsWith(',') ||
		integerPicture.endsWith(',') ||
		integerPicture.includes(',,') ||
		/0#/.test(match[2] ?? '')
	)
		return formulaFailure('unsupported', `FORMAT picture "${picture}"`);
	const fractionPicture = match[2] ?? '',
		percent = match[3] === '%',
		grouped = integerPicture.includes(',');
	const minInteger = (integerPicture.match(/0/g) ?? []).length;
	const minFraction = (fractionPicture.match(/0/g) ?? []).length;
	const maxFraction = fractionPicture.length;
	const scaled = percent ? value.value * 100 : value.value;
	// Half away from zero on the decimal digits, as a spreadsheet picture does.
	const factor = 10 ** maxFraction;
	const rounded = (Math.sign(scaled) * Math.round(Math.abs(scaled) * factor + 1e-9)) / factor;
	const [digits = '', decimals = ''] = Math.abs(rounded).toFixed(maxFraction).split('.');
	let whole = digits.replace(/^0+/, '').padStart(minInteger, '0');
	if (grouped) whole = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
	let fraction = decimals;
	while (fraction.length > minFraction && fraction.endsWith('0')) fraction = fraction.slice(0, -1);
	const negative = rounded < 0;
	return `${negative ? '-' : ''}${whole}${fraction ? `.${fraction}` : ''}${percent ? '%' : ''}`;
}

/** DATE(year, month, day) as an OLE date serial (days since 1899-12-30). */
export function visioDateSerial(year: number, month: number, day: number): number {
	if (![year, month, day].every(Number.isInteger))
		return formulaFailure('value', 'DATE requires whole numbers');
	if (year < 1 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31)
		return formulaFailure('value', 'DATE is outside the calendar');
	const date = new Date(0);
	date.setUTCFullYear(year, month - 1, day);
	if (date.getUTCMonth() !== month - 1) return formulaFailure('value', 'DATE is not a real day');
	return (date.getTime() - Date.UTC(1899, 11, 30)) / 86_400_000;
}

export interface ShapeFrame {
	pinX: number;
	pinY: number;
	locPinX: number;
	locPinY: number;
	/** Radians. */
	angle: number;
}

/** ANGLETOLOC: an angle seen from the shape's parent, seen from the shape itself. */
export const angleToLocal = (frame: ShapeFrame, angle: number): number => angle - frame.angle;

/** LOCTOPAR: a point in the shape's local frame, expressed in its parent's frame. */
export function localToParent(frame: ShapeFrame, x: number, y: number): [number, number] {
	const dx = x - frame.locPinX,
		dy = y - frame.locPinY,
		cos = Math.cos(frame.angle),
		sin = Math.sin(frame.angle);
	return [frame.pinX + dx * cos - dy * sin, frame.pinY + dx * sin + dy * cos];
}
