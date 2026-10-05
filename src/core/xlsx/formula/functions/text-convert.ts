import { formatValue } from '../../numfmt/index.js';
import { parseNumberText } from '../text-number.js';
import { ERR, fail, isError } from '../values.js';
import { bool, int, num, optNum, scalar, spec, str } from './helpers.js';
import { roundTo } from './math.js';
import type { FunctionSpec } from './types.js';

const C = 'Text';

/** Windows-1252 characters for codes 128-159 (CHAR and CODE use the ANSI code page). */
const CP1252: Record<number, number> = {
	128: 0x20ac,
	130: 0x201a,
	131: 0x0192,
	132: 0x201e,
	133: 0x2026,
	134: 0x2020,
	135: 0x2021,
	136: 0x02c6,
	137: 0x2030,
	138: 0x0160,
	139: 0x2039,
	140: 0x0152,
	142: 0x017d,
	145: 0x2018,
	146: 0x2019,
	147: 0x201c,
	148: 0x201d,
	149: 0x2022,
	150: 0x2013,
	151: 0x2014,
	152: 0x02dc,
	153: 0x2122,
	154: 0x0161,
	155: 0x203a,
	156: 0x0153,
	158: 0x017e,
	159: 0x0178,
};

const CP1252_REVERSE = new Map(Object.entries(CP1252).map(([k, v]) => [v, Number(k)]));

/** Formats with thousands separators and fixed decimals (FIXED, DOLLAR). */
function fixed(value: number, decimals: number, commas: boolean): string {
	const d = Math.trunc(decimals);
	const rounded = roundTo(value, d, 'half');
	const abs = Math.abs(rounded);
	let text = abs.toFixed(Math.max(0, Math.min(d, 99)));
	if (commas) {
		const [int = '', frac] = text.split('.');
		text = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (frac !== undefined ? `.${frac}` : '');
	}
	return rounded < 0 ? `-${text}` : text;
}

/** Conversions between numbers, codes and text. */
export const TEXT_CONVERT: FunctionSpec[] = [
	spec(
		'TEXT',
		C,
		'TEXT(value, format_text)',
		'Formats a value with a number format.',
		2,
		2,
		(args, ctx) => {
			let value = scalar(args[0]);
			if (isError(value)) return value;
			if (typeof value === 'string') value = parseNumberText(value) ?? value;
			if (typeof value === 'boolean') value = value ? 'TRUE' : 'FALSE';
			return formatValue(value, str(args[1]), { date1904: ctx.date1904 }).text;
		},
	),
	spec('VALUE', C, 'VALUE(text)', 'Converts text to a number.', 1, 1, (args) => {
		const v = scalar(args[0]);
		if (typeof v === 'number') return v;
		if (v === null) return 0;
		if (typeof v === 'boolean') fail(ERR.VALUE);
		const parsed = parseNumberText(str(v));
		return parsed === undefined ? fail(ERR.VALUE) : parsed;
	}),
	spec(
		'NUMBERVALUE',
		C,
		'NUMBERVALUE(text, [decimal_separator], [group_separator])',
		'Converts text to a number with given separators.',
		1,
		3,
		(args) => {
			const decimal = (args.length > 1 ? str(args[1]) : '.').charAt(0) || '.';
			const group = (args.length > 2 ? str(args[2]) : ',').charAt(0) || ',';
			let text = str(args[0]).replace(/\s+/g, '');
			if (text === '') return 0;
			let percent = 0;
			while (text.endsWith('%')) {
				percent++;
				text = text.slice(0, -1);
			}
			const decimalAt = text.indexOf(decimal);
			if (decimalAt >= 0 && text.slice(decimalAt + 1).includes(group)) fail(ERR.VALUE);
			text = text.split(group).join('');
			if (decimal !== '.') text = text.replace(decimal, '.');
			const value = Number(text);
			if (!/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(text) || !Number.isFinite(value))
				fail(ERR.VALUE);
			return value / Math.pow(100, percent);
		},
	),
	spec(
		'FIXED',
		C,
		'FIXED(number, [decimals], [no_commas])',
		'Formats a number as text with fixed decimals.',
		1,
		3,
		(args) => fixed(num(args[0]), optNum(args, 1, 2), !(args.length > 2 && bool(args[2]))),
	),
	spec(
		'DOLLAR',
		C,
		'DOLLAR(number, [decimals])',
		'Formats a number as currency text.',
		1,
		2,
		(args) => {
			const value = num(args[0]);
			const text = fixed(Math.abs(value), optNum(args, 1, 2), true);
			return value < 0 && text.replace(/[0.,]/g, '') !== '' ? `($${text})` : `$${text}`;
		},
	),
	spec('CHAR', C, 'CHAR(number)', 'The character for a code (Windows-1252).', 1, 1, (args) => {
		const n = int(args[0]);
		if (n < 1 || n > 255) fail(ERR.VALUE);
		return String.fromCharCode(CP1252[n] ?? n);
	}),
	spec('CODE', C, 'CODE(text)', 'The code of the first character (Windows-1252).', 1, 1, (args) => {
		const text = str(args[0]);
		if (text === '') fail(ERR.VALUE);
		const code = text.charCodeAt(0);
		return CP1252_REVERSE.get(code) ?? (code <= 255 ? code : 63);
	}),
	spec('UNICHAR', C, 'UNICHAR(number)', 'The Unicode character for a code point.', 1, 1, (args) => {
		const n = int(args[0]);
		if (n < 1 || n > 0x10ffff) fail(ERR.VALUE);
		return String.fromCodePoint(n);
	}),
	spec('UNICODE', C, 'UNICODE(text)', 'The code point of the first character.', 1, 1, (args) => {
		const text = str(args[0]);
		if (text === '') fail(ERR.VALUE);
		return text.codePointAt(0) ?? fail(ERR.VALUE);
	}),
	spec('T', C, 'T(value)', 'The text of a value, or empty text.', 1, 1, (args) => {
		const v = scalar(args[0]);
		if (isError(v)) return v;
		return typeof v === 'string' ? v : '';
	}),
];
