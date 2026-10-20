// The decimal and thousands separators of a CSV import (the Text Import Wizard's Advanced
// settings): validation, and the rewrite to the `.` and `,` that cell typing reads.

export interface Separators {
	decimal: string;
	thousands: string;
}

/** The decimal and thousands separators, validated; the thousands one follows the decimal one. */
export function separatorsOf(decimal = '.', thousands = decimal === ',' ? '.' : ','): Separators {
	for (const separator of [decimal, thousands])
		if (separator.length !== 1 || /[\d+-]/.test(separator))
			throw new RangeError(`"${separator}" cannot be a decimal or thousands separator.`);
	if (decimal === thousands)
		throw new RangeError('The decimal and thousands separators must differ.');
	return { decimal, thousands };
}

/**
 * Rewrites the configured separators as `.` and `,` (and any other `.` or `,` as a character no
 * number contains), or `undefined` for the default separators. Only a thousands separator
 * between digits groups, so the spaces of `1 234,56 €` around the number stay.
 */
export function separatorNormalizer({
	decimal,
	thousands,
}: Separators): ((text: string) => string) | undefined {
	if (decimal === '.' && thousands === ',') return undefined;
	const escape = (char: string): string => char.replace(/[\\^$.*+?()[\]{}|]/g, '\\$&');
	const pattern = new RegExp(`(${escape(decimal)})|(\\d)${escape(thousands)}(?=\\d)|[.,]`, 'g');
	const swap = (_: string, point?: string, digit?: string): string => {
		if (point !== undefined) return '.';
		if (digit !== undefined) return `${digit},`;
		return '\u0001';
	};
	return (text) => text.replace(pattern, swap);
}
