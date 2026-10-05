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

export type DigitChar = '0' | '#' | '?';

export type DatePart =
	| 'y'
	| 'yyyy'
	| 'e'
	/** Japanese era name (`g`, `gg`, `ggg`): empty outside Japanese locales. */
	| 'g'
	| 'b'
	| 'bbbb'
	| 'm'
	| 'mm'
	| 'mmm'
	| 'mmmm'
	| 'mmmmm'
	| 'd'
	| 'dd'
	| 'ddd'
	| 'dddd'
	| 'h'
	| 'hh'
	| 'min'
	| 'mmin'
	| 's'
	| 'ss';

/** One token of a format section. */
export type Token =
	| { t: 'lit'; v: string }
	| { t: 'digit'; c: DigitChar }
	/** A literal digit 1-9 (a fixed fraction denominator when it follows `/`). */
	| { t: 'num'; v: string }
	| { t: 'point' }
	| { t: 'comma' }
	| { t: 'percent' }
	| { t: 'exp'; plus: boolean }
	| { t: 'slash' }
	| { t: 'text' }
	| { t: 'general' }
	| { t: 'date'; part: DatePart }
	| { t: 'elapsed'; unit: 'h' | 'm' | 's'; width: number }
	| { t: 'ampm'; kind: 'AM/PM' | 'A/P' | 'a/p' }
	| { t: 'subsec'; digits: number };

export type ConditionOp = '<' | '<=' | '>' | '>=' | '=' | '<>';

export interface Condition {
	op: ConditionOp;
	value: number;
}

export type SectionKind = 'number' | 'date' | 'text' | 'general';

export interface Section {
	tokens: Token[];
	kind: SectionKind;
	color?: string;
	condition?: Condition;
}

export interface CompiledFormat {
	/** Number sections (at most 3). */
	numeric: Section[];
	/** Text section (the 4th section, or the only section when it contains `@`). */
	text?: Section;
	/** True when any section carries a condition. */
	conditional: boolean;
	isDate: boolean;
}
