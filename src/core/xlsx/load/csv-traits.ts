// The traits of a typed CSV field that the import's conversion switches and warnings act on:
// leading zeros, long numbers, exponents and letters-and-digits dates, read from the field's text.

// Bit `1 << i` of a traits mask is the reader's conversion kind `i` (`KINDS` in csv-import).
export const LEADING_ZEROS = 1;
export const LONG_NUMBER = 2;
export const SCIENTIFIC = 4;
export const LETTER_DATE = 8;
export const TIME_ZONE = 16;
/** Not a trait of the text: a date-column value that the General fallback read as a date. */
export const DATE_FALLBACK = 32;

const isDigit = (code: number): boolean => code >= 48 && code <= 57;
const isLetter = (code: number): boolean => (code | 32) >= 97 && (code | 32) <= 122;

/**
 * Leading zeros (in the integer part only: `0,123` has one, `.05` none), 16+ significant digits
 * and an exponent, from a field typed as a number with these separators (char codes).
 */
export function numberTraits(text: string, decimal: number, thousands: number): number {
	let traits = 0;
	let digits = 0;
	let first = true;
	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i);
		if (isDigit(code)) {
			if (first && code === 48) {
				const next = text.charCodeAt(i + 1);
				if (isDigit(next) || (next === thousands && isDigit(text.charCodeAt(i + 2))))
					traits |= LEADING_ZEROS;
			}
			first = false;
			if (digits > 0 || code !== 48) digits++;
		} else if (code === decimal) first = false;
		else if (code === 69 || code === 101) {
			traits |= SCIENTIFIC;
			break;
		}
	}
	return digits > 15 ? traits | LONG_NUMBER : traits;
}

/** The significant digits of a number's text: no sign, separators, exponent or outer zeros. */
function significantDigits(text: string): string {
	let digits = '';
	for (const char of text) {
		if (char === 'e' || char === 'E') break;
		if (char >= '0' && char <= '9' && (digits !== '' || char !== '0')) digits += char;
	}
	let end = digits.length;
	while (end > 0 && digits[end - 1] === '0') end--;
	return digits.slice(0, end);
}

/** The number keeps every significant digit of the field (`0.30000000000000004` does). */
export const keepsDigits = (field: string, value: number): boolean =>
	significantDigits(field) === significantDigits(String(Math.abs(value)));

/**
 * A date typed from letters followed directly by digits (`JAN1`, `SEPT2`), checked on the
 * trimmed text as `parseCellInput` reads it (` JAN1` is one too).
 */
export function dateTraits(field: string): number {
	const text = field.trimStart();
	let i = 0;
	while (isLetter(text.charCodeAt(i))) i++;
	return i > 0 && isDigit(text.charCodeAt(i)) ? LETTER_DATE : 0;
}
