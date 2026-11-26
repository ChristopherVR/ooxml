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
