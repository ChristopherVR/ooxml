/**
 * Built-in number formats (`numFmtId` values a workbook may use without declaring them), as Excel
 * shows them in the en-US locale. Ids 23-36 and 50+ are locale specific (East Asian) and omitted.
 */
export const BUILTIN_NUMBER_FORMATS: Readonly<Record<number, string>> = Object.freeze({
	0: 'General',
	1: '0',
	2: '0.00',
	3: '#,##0',
	4: '#,##0.00',
	5: '"$"#,##0_);\\("$"#,##0\\)',
	6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
	7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
	8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
	9: '0%',
	10: '0.00%',
	11: '0.00E+00',
	12: '# ?/?',
	13: '# ??/??',
	14: 'm/d/yyyy',
	15: 'd-mmm-yy',
	16: 'd-mmm',
	17: 'mmm-yy',
	18: 'h:mm AM/PM',
	19: 'h:mm:ss AM/PM',
	20: 'h:mm',
	21: 'h:mm:ss',
	22: 'm/d/yyyy h:mm',
	// Excel 16 en-US maps these codes to ids 37-40 (verified by saving them from Excel).
	37: '#,##0_);(#,##0)',
	38: '#,##0_);[Red](#,##0)',
	39: '#,##0.00_);(#,##0.00)',
	40: '#,##0.00_);[Red](#,##0.00)',
	41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
	42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
	43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
	44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
	45: 'mm:ss',
	46: '[h]:mm:ss',
	// ECMA-376 lists `mmss.0` (a known erratum); Excel shows and applies `mm:ss.0`.
	47: 'mm:ss.0',
	48: '##0.0E+0',
	49: '@',
});

/** Spellings other producers use for the same built-in formats (ECMA-376 lists these). */
const ALIASES: Readonly<Record<string, number>> = {
	general: 0,
	'mm-dd-yy': 14,
	'm/d/yy': 14,
	'm/d/yy h:mm': 22,
	'h:mm am/pm': 18,
	'h:mm:ss am/pm': 19,
	'#,##0 ;(#,##0)': 37,
	'#,##0 ;[red](#,##0)': 38,
	'#,##0.00;(#,##0.00)': 39,
	'#,##0.00;[red](#,##0.00)': 40,
	'#,##0 ;\\(#,##0\\)': 37,
	'#,##0 ;[red]\\(#,##0\\)': 38,
	'#,##0.00;\\(#,##0.00\\)': 39,
	'#,##0.00;[red]\\(#,##0.00\\)': 40,
	'mmss.0': 47,
};

let reverse: Map<string, number> | undefined;

/** The built-in id of a format code, or `undefined` when it has to be declared in `numFmts`. */
export function builtinFormatId(format: string): number | undefined {
	if (!reverse) {
		reverse = new Map();
		for (const [id, code] of Object.entries(BUILTIN_NUMBER_FORMATS)) reverse.set(code, Number(id));
	}
	return reverse.get(format) ?? ALIASES[format.toLowerCase()];
}

export interface FormatPreset {
	id: string;
	label: string;
	format: string;
}

/** The Home ribbon's number format drop-down. */
export const FORMAT_PRESETS: FormatPreset[] = [
	{ id: 'general', label: 'General', format: 'General' },
	{ id: 'number', label: 'Number', format: '0.00' },
	{ id: 'currency', label: 'Currency', format: '"$"#,##0.00' },
	{ id: 'accounting', label: 'Accounting', format: BUILTIN_NUMBER_FORMATS[44] ?? '' },
	{ id: 'shortDate', label: 'Short Date', format: 'm/d/yyyy' },
	{ id: 'longDate', label: 'Long Date', format: 'dddd, mmmm d, yyyy' },
	{ id: 'time', label: 'Time', format: 'h:mm:ss AM/PM' },
	{ id: 'percentage', label: 'Percentage', format: '0.00%' },
	{ id: 'fraction', label: 'Fraction', format: '# ?/?' },
	{ id: 'scientific', label: 'Scientific', format: '0.00E+00' },
	{ id: 'text', label: 'Text', format: '@' },
];
