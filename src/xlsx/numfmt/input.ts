import { type CellValue, cellError, isErrorCode } from '../model.js';
import { parseDateTimeInput } from './input-date.js';
import type { FormatOptions } from './types.js';

/** What typing `text` into a cell produces. */
export interface ParsedInput {
	value: CellValue;
	/** Number format Excel applies automatically (`0%`, `#,##0.00`, `m/d/yyyy`, ...). */
	numFmt?: string;
	/** The formula without its leading `=`. */
	formula?: string;
}

const NUMBER_RE = /^(?:(\d{1,3}(?:,\d{3})+)|(\d*))(?:\.(\d*))?$/;

interface NumberInput {
	value: number;
	numFmt?: string;
}

function currencyFormat(symbol: string, decimals: boolean): string {
	const body = decimals ? '#,##0.00' : '#,##0';
	if (symbol === '$') return `$${body}_);[Red]($${body})`;
	return `${symbol}${body}`;
}

/** Numbers with sign, parentheses, thousands separators, currency, percent or exponent. */
export function parseNumberInput(text: string): NumberInput | undefined {
	let s = text.trim();
	let negative = false;
	if (s.startsWith('(') && s.endsWith(')')) {
		negative = true;
		s = s.slice(1, -1).trim();
	}
	const sign = (): void => {
		if (s.startsWith('-') || s.startsWith('+')) {
			if (s.startsWith('-')) negative = !negative;
			s = s.slice(1).trim();
		}
	};
	sign();
	let currency = '';
	const lead = /^([$€£])\s*/.exec(s);
	if (lead) {
		currency = lead[1] ?? '';
		s = s.slice(lead[0].length);
		sign();
	} else {
		const trail = /\s*([€£])$/.exec(s);
		if (trail) {
			currency = trail[1] ?? '';
			s = s.slice(0, -trail[0].length);
		}
	}
	let percent = false;
	if (!currency && s.endsWith('%')) {
		percent = true;
		s = s.slice(0, -1).trim();
	}
	let exponent = '';
	const exp = /[eE]([-+]?\d+)$/.exec(s);
	if (exp && !currency && !percent) {
		exponent = exp[1] ?? '';
		s = s.slice(0, -exp[0].length);
	}
	const m = NUMBER_RE.exec(s);
	if (!m || !/\d/.test(s)) return undefined;
	const grouped = m[1] !== undefined;
	if (grouped && exponent) return undefined;
	const decimals = m[3] !== undefined && m[3] !== '';
	const core = `${(m[1] ?? m[2] ?? '').replace(/,/g, '') || '0'}.${m[3] ?? ''}0`;
	let value = Number(`${core}e${exponent ? Number(exponent) : 0}`);
	if (percent) value = Number(`${core}e-2`);
	if (!Number.isFinite(value)) return undefined;
	if (negative && value !== 0) value = -value;
	if (exponent) return { value, numFmt: '0.00E+00' };
	if (percent) return { value, numFmt: decimals ? '0.00%' : '0%' };
	if (currency) return { value, numFmt: currencyFormat(currency, decimals) };
	if (grouped) return { value, numFmt: decimals ? '#,##0.00' : '#,##0' };
	return { value };
}

function parseFractionInput(text: string): NumberInput | undefined {
	const m = /^([-+]?)(\d+)\s+(\d+)\/(\d+)$/.exec(text.trim());
	if (!m) return undefined;
	const den = Number(m[4]);
	const num = Number(m[3]);
	if (den === 0 || num >= den) return undefined;
	const magnitude = Number(m[2]) + num / den;
	return {
		value: m[1] === '-' ? -magnitude : magnitude,
		numFmt: den < 10 ? '# ?/?' : '# ??/??',
	};
}

/**
 * Interprets text typed into a cell like Excel's en-US entry: `=` starts a formula, a leading
 * apostrophe forces text, then booleans, error literals, numbers, fractions, dates and times.
 */
export function parseCellInput(text: string, options: FormatOptions = {}): ParsedInput {
	if (text === '') return { value: null };
	if (text.startsWith('=') && text.length > 1) return { value: null, formula: text.slice(1) };
	if (text.startsWith("'")) return { value: text.slice(1) };
	const trimmed = text.trim();
	if (trimmed === '') return { value: text };
	const upper = trimmed.toUpperCase();
	if (upper === 'TRUE' || upper === 'FALSE') return { value: upper === 'TRUE' };
	if (isErrorCode(upper)) return { value: cellError(upper) };
	const parsed =
		parseNumberInput(trimmed) ??
		parseFractionInput(trimmed) ??
		parseDateTimeInput(trimmed, options.date1904 ?? false);
	if (!parsed) return { value: text };
	return parsed.numFmt ? { value: parsed.value, numFmt: parsed.numFmt } : { value: parsed.value };
}
