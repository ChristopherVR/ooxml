import { type Digits, plain, roundDecimals, roundSignificant, toDigits } from './decimal.js';

/** Width Excel's General format fits a number into (sign excluded). */
const GENERAL_WIDTH = 11;

function scientific(v: Digits, width = GENERAL_WIDTH): { text: string; sig: number } {
	let exp = v.point - 1;
	let expDigits = Math.max(2, String(Math.abs(exp)).length);
	let r = roundSignificant(v, Math.max(1, width - 3 - expDigits));
	if (r.point - 1 !== exp) {
		exp = r.point - 1;
		expDigits = Math.max(2, String(Math.abs(exp)).length);
		r = roundSignificant(v, Math.max(1, width - 3 - expDigits));
	}
	const rest = r.d.slice(1);
	const mantissa = rest ? `${r.d.charAt(0)}.${rest}` : r.d.charAt(0);
	const expText = String(Math.abs(exp)).padStart(2, '0');
	return { text: `${mantissa}E${exp < 0 ? '-' : '+'}${expText}`, sig: r.d.length };
}

/** Formats a non-negative digit value the way the General format does. */
function generalDigits(v: Digits, width: number): string {
	if (!v.d) return '0';
	const exp = v.point - 1;
	if (exp >= width) return scientific(v, width).text;
	if (exp >= 0) {
		const r = roundDecimals(v, Math.max(0, width - 2 - exp));
		if (r.point > width) return scientific(v, width).text;
		return plain(r);
	}
	const fixed = roundDecimals(v, Math.max(0, width - 2));
	const sci = scientific(v, width);
	if (fixed.d && fixed.d.length >= sci.sig) return plain(fixed);
	return sci.text;
}

/**
 * Excel's General number format for a cell wide enough for 11 characters: up to 11 digits,
 * switching to scientific notation (`1.23457E+11`) when that keeps more precision. A smaller
 * `width` (characters, sign excluded, at least 1) fits narrower columns the same way: fewer
 * decimals first, then scientific notation with fewer significant digits.
 */
export function formatGeneral(value: number, width = GENERAL_WIDTH): string {
	if (!Number.isFinite(value)) return '#NUM!';
	const text = generalDigits(
		toDigits(value),
		Math.max(1, Math.min(GENERAL_WIDTH, Math.floor(width))),
	);
	return value < 0 && text !== '0' ? `-${text}` : text;
}
