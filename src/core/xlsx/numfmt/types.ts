/** The display result of formatting a value with a number format code. */
export interface FormattedValue {
	text: string;
	/** `#RRGGBB` from a colour token (`[Red]`, `[Color10]`) of the section that was used. */
	color?: string;
}

export interface FormatOptions {
	/** Workbook uses the 1904 date system (serial 0 is 1904-01-01). */
	date1904?: boolean;
	/**
	 * Room the cell has for a number, like a column width. Without `measure` it counts characters;
	 * with `measure` it is in the units `measure` returns (for example CSS pixels). A General
	 * number that does not fit drops decimals, then switches to shorter scientific notation; any
	 * number (or date) that still does not fit shows `#` repeated across the width, as Excel does.
	 * Text, booleans and errors are never fitted. Absent: no fitting (the 11-character General).
	 */
	width?: number;
	/** Measures a candidate text in the same units as `width`. */
	measure?: (text: string) => number;
}

export type {
	Condition,
	DatePart,
	DigitChar,
	Section,
	Token,
} from '../../text/number-format/types';
