import { type CellValue, isCellError } from '../model.js';
import { chooseSection, type CompiledSection, compileFormat } from './compile.js';
import { renderDate } from './date.js';
import { renderFraction } from './fraction.js';
import { formatGeneral } from './general.js';
import { renderNumber } from './number.js';
import type { FormatOptions, FormattedValue, Token } from './types.js';

/** What Excel shows when a value does not fit its format (negative date, no matching section). */
export const OVERFLOW_TEXT = '########';

const result = (text: string, section?: CompiledSection): FormattedValue =>
	section?.color ? { text, color: section.color } : { text };

function tokenSource(tok: Token): string {
	switch (tok.t) {
		case 'lit':
			return tok.v;
		case 'digit':
			return tok.c;
		case 'num':
			return tok.v;
		case 'point':
			return '.';
		case 'comma':
			return ',';
		case 'percent':
			return '%';
		case 'slash':
			return '/';
		case 'exp':
			return tok.plus ? 'E+' : 'E-';
		case 'general':
			return 'General';
		case 'date':
			return tok.part === 'min' ? 'm' : tok.part === 'mmin' ? 'mm' : tok.part;
		case 'ampm':
			return tok.kind;
		case 'subsec':
			return `.${'0'.repeat(tok.digits)}`;
		case 'elapsed':
			return tok.unit.repeat(tok.width);
		default:
			return '';
	}
}

function formatText(text: string, section: CompiledSection | undefined): FormattedValue {
	if (!section) return { text };
	let out = '';
	for (const tok of section.tokens) out += tok.t === 'text' ? text : tokenSource(tok);
	return result(out, section);
}

/** Renders `|v|` with a general/number/fraction section; `zero` when it displays as zero. */
function renderAbs(
	v: number,
	section: CompiledSection,
	generalWidth?: number,
): { text: string; zero: boolean } {
	if (section.fraction) return renderFraction(v, section.fraction);
	if (section.number) return renderNumber(v, section.number);
	const general = formatGeneral(Math.abs(v), generalWidth);
	let text = '';
	for (const tok of section.tokens) {
		text += tok.t === 'general' || tok.t === 'text' ? general : tokenSource(tok);
	}
	return { text, zero: general === '0' };
}

/** `#` repeated across `width` (at least one), Excel's display for a number that does not fit. */
function hashFill(options: FormatOptions): string {
	const width = options.width ?? 0;
	const one = options.measure ? options.measure('#') : 1;
	return '#'.repeat(Math.max(1, Math.floor(one > 0 ? width / one : width)));
}

function fitNumber(value: number, format: string, options: FormatOptions): FormattedValue {
	const date1904 = options.date1904 ?? false;
	const width = options.width ?? 0;
	const fits = (text: string): boolean =>
		(options.measure ? options.measure(text) : text.length) <= width;
	const first = formatNumber(value, format, date1904);
	if (fits(first.text) || first.text === OVERFLOW_TEXT) return first;
	const choice = chooseSection(compileFormat(format), value);
	const general =
		choice &&
		!choice.section.number &&
		!choice.section.fraction &&
		choice.section.tokens.some((t) => t.t === 'general');
	if (general) {
		for (let w = 10; w >= 1; w--) {
			const next = formatNumber(value, format, date1904, w);
			if (fits(next.text)) return next;
		}
	}
	const fill = hashFill(options);
	return first.color ? { text: fill, color: first.color } : { text: fill };
}

function formatNumber(
	value: number,
	format: string,
	date1904: boolean,
	generalWidth?: number,
): FormattedValue {
	if (!Number.isFinite(value)) return { text: '#NUM!' };
	const compiled = compileFormat(format);
	const choice = chooseSection(compiled, value);
	if (!choice) return { text: OVERFLOW_TEXT };
	const { section, abs } = choice;
	if (section.kind === 'date') {
		const serial = abs ? Math.abs(value) : value;
		const text = renderDate(serial, section.tokens, date1904);
		return text === undefined ? { text: OVERFLOW_TEXT } : result(text, section);
	}
	const { text, zero } = renderAbs(value, section, generalWidth);
	const minus = value < 0 && !abs && !zero;
	return result(minus ? `-${text}` : text, section);
}

/**
 * Formats a cell value with an Excel number format code (`#,##0.00`, `m/d/yyyy`, `[Red]0;0`, ...).
 * Fill characters (`*x`) are dropped and padding (`_x`) becomes one space. With
 * `options.width`, numbers are fitted to the cell the way Excel does (see {@link FormatOptions}).
 */
export function formatValue(
	value: CellValue,
	format: string,
	options: FormatOptions = {},
): FormattedValue {
	if (value === null) return { text: '' };
	if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE' };
	if (typeof value === 'string') return formatText(value, compileFormat(format).text);
	if (isCellError(value)) return { text: value.error };
	if (options.width !== undefined) return fitNumber(value, format, options);
	return formatNumber(value, format, options.date1904 ?? false);
}

/** True when the format shows numbers as dates or times. */
export function isDateFormat(format: string): boolean {
	return compileFormat(format).isDate;
}
