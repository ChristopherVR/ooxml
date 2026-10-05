// Number tab categories: the format code each category's options produce and, for an existing
// code, the category and options that reproduce it (Custom when none does).
export type CategoryId =
	| 'general'
	| 'number'
	| 'currency'
	| 'accounting'
	| 'date'
	| 'time'
	| 'percentage'
	| 'fraction'
	| 'scientific'
	| 'text'
	| 'special'
	| 'custom';

export const CATEGORIES: ReadonlyArray<readonly [CategoryId, string]> = [
	['general', 'General'],
	['number', 'Number'],
	['currency', 'Currency'],
	['accounting', 'Accounting'],
	['date', 'Date'],
	['time', 'Time'],
	['percentage', 'Percentage'],
	['fraction', 'Fraction'],
	['scientific', 'Scientific'],
	['text', 'Text'],
	['special', 'Special'],
	['custom', 'Custom'],
];

export interface NumberOptions {
	decimals: number;
	thousands: boolean;
	/** Index into SYMBOLS. */
	symbol: number;
	/** Index into the four negative number styles. */
	negative: number;
	/** The chosen code of a list category (date, time, fraction, special) or the custom code. */
	code: string;
}

/** Currency symbols: [label, quoted code prefix]. */
export const SYMBOLS: ReadonlyArray<readonly [string, string]> = [
	['None', ''],
	['$', '"$"'],
	['€', '"€"'],
	['£', '"£"'],
	['¥', '"¥"'],
];

export const DATE_CODES = [
	'm/d/yyyy',
	'dddd, mmmm d, yyyy',
	'm/d',
	'm/d/yy',
	'mm/dd/yy',
	'd-mmm',
	'd-mmm-yy',
	'dd-mmm-yy',
	'mmm-yy',
	'mmmm-yy',
	'mmmm d, yyyy',
	'm/d/yy h:mm AM/PM',
	'm/d/yy h:mm',
	'mmmmm',
	'mmmmm-yy',
	'd-mmm-yyyy',
	'yyyy-mm-dd',
] as const;

export const TIME_CODES = [
	'h:mm:ss AM/PM',
	'h:mm',
	'h:mm AM/PM',
	'h:mm:ss',
	'mm:ss',
	'mm:ss.0',
	'[h]:mm:ss',
	'm/d/yy h:mm AM/PM',
	'm/d/yy h:mm',
] as const;

export const FRACTION_CODES: ReadonlyArray<readonly [string, string]> = [
	['# ?/?', 'Up to one digit (1/4)'],
	['# ??/??', 'Up to two digits (21/25)'],
	['# ???/???', 'Up to three digits (312/943)'],
	['# ?/2', 'As halves (1/2)'],
	['# ?/4', 'As quarters (2/4)'],
	['# ?/8', 'As eighths (4/8)'],
	['# ??/16', 'As sixteenths (8/16)'],
	['# ?/10', 'As tenths (3/10)'],
	['# ??/100', 'As hundredths (30/100)'],
];

export const SPECIAL_CODES: ReadonlyArray<readonly [string, string]> = [
	['00000', 'Zip Code'],
	['00000-0000', 'Zip Code + 4'],
	['[<=9999999]###-####;(###) ###-####', 'Phone Number'],
	['000-00-0000', 'Social Security Number'],
];

const decimals = (n: number): string => (n > 0 ? `.${'0'.repeat(n)}` : '');

/** The four negative number styles of a positive code: -1234, red, (1234), red (1234). */
export function negativeCodes(base: string): string[] {
	return [base, `${base};[Red]${base}`, `${base}_);(${base})`, `${base}_);[Red](${base})`];
}

export function buildCode(category: CategoryId, o: NumberOptions): string {
	const d = decimals(o.decimals);
	switch (category) {
		case 'general':
			return 'General';
		case 'number':
			return negativeCodes(`${o.thousands ? '#,##0' : '0'}${d}`)[o.negative] ?? '0';
		case 'currency':
			return negativeCodes(`${SYMBOLS[o.symbol]?.[1] ?? ''}#,##0${d}`)[o.negative] ?? '#,##0';
		case 'accounting': {
			const s = SYMBOLS[o.symbol]?.[1] ?? '';
			const sp = s ? `${s}* ` : '* ';
			return `_(${sp}#,##0${d}_);_(${sp}\\(#,##0${d}\\);_(${sp}"-"${'?'.repeat(o.decimals)}_);_(@_)`;
		}
		case 'percentage':
			return `0${d}%`;
		case 'scientific':
			return `0${d}E+00`;
		case 'text':
			return '@';
		default:
			return o.code;
	}
}

export const DEFAULT_OPTIONS: NumberOptions = {
	decimals: 2,
	thousands: false,
	symbol: 1,
	negative: 0,
	code: '',
};

/** The category and options for an existing code (Custom when no option set reproduces it). */
export function detectCategory(code: string): { category: CategoryId; options: NumberOptions } {
	const base = { ...DEFAULT_OPTIONS, code };
	if (/^general$/i.test(code)) return { category: 'general', options: base };
	if (code === '@') return { category: 'text', options: base };
	const lists: [CategoryId, readonly string[]][] = [
		['special', SPECIAL_CODES.map(([c]) => c)],
		['date', DATE_CODES.filter((c) => !TIME_CODES.includes(c as (typeof TIME_CODES)[number]))],
		['time', TIME_CODES],
		['fraction', FRACTION_CODES.map(([c]) => c)],
	];
	for (const [category, codes] of lists)
		if (codes.includes(code)) return { category, options: base };
	for (const category of ['number', 'currency', 'accounting', 'percentage', 'scientific'] as const)
		for (let dec = 0; dec <= 10; dec++)
			for (const thousands of [false, true])
				for (let symbol = 0; symbol < SYMBOLS.length; symbol++)
					for (let negative = 0; negative < 4; negative++) {
						const options = { decimals: dec, thousands, symbol, negative, code };
						if (buildCode(category, options) === code) return { category, options };
					}
	return { category: 'custom', options: base };
}
