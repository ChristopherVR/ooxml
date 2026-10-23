/**
 * Numbers typed or pasted by a person, read and written in their locale: a `de-DE` author types
 * `1.234,5`, an `en-US` author `1,234.5`, and both mean the same value. Used by the chart data
 * sheet; DOM-free and format-neutral.
 *
 * @module text/locale-number
 */
import { decimalSeparatorForLanguage } from './decimal-tab';

/** Spaces a locale may group digits with (`fr-FR` writes `1 234,5` with a narrow no-break space). */
const GROUP_SPACES = /[\s  ]/g;

/**
 * Reads `raw` as a number in `language`'s convention, or `null` when it is not one. The locale's
 * decimal separator is the decimal point; the other of `.` and `,` is accepted only as a digit
 * group separator in well-formed groups of three (`1.234.567`). A lone `.` in a comma-decimal
 * locale is still read as a decimal point (`1.5`), the common case of a value pasted from an
 * English source. Exponents (`1e3`) and a leading sign are accepted; empty text is not a number.
 */
export function parseLocaleNumber(raw: string, language?: string): number | null {
	const decimal = decimalSeparatorForLanguage(language);
	const group = decimal === ',' ? '.' : ',';
	let text = raw.trim().replace(GROUP_SPACES, '');
	if (text === '') {
		return null;
	}
	const groups = new RegExp(
		`^[+-]?\\d{1,3}(?:\\${group}\\d{3})+(?:\\${decimal}\\d*)?(?:e[+-]?\\d+)?$`,
		'i',
	);
	if (groups.test(text)) {
		text = text.split(group).join('');
	} else if (text.includes(group) && !(decimal === ',' && !text.includes(','))) {
		return null;
	}
	if (decimal !== '.') {
		text = text.replace(decimal, '.');
	}
	if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(text)) {
		return null;
	}
	const value = Number(text);
	return Number.isFinite(value) ? value : null;
}

/**
 * Writes `value` for editing in `language`'s convention: its decimal separator, no digit
 * grouping and every significant digit, so reading the text back with {@link parseLocaleNumber}
 * gives the same number.
 */
export function formatLocaleNumber(value: number, language?: string): string {
	if (!Number.isFinite(value)) {
		return '';
	}
	const plain = String(value);
	if (/e/i.test(plain)) {
		return plain;
	}
	const decimal = decimalSeparatorForLanguage(language);
	return decimal === '.' ? plain : plain.replace('.', decimal);
}
