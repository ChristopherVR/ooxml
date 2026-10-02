import { toText } from '../coerce.js';
import { ERR, fail, type Value } from '../values.js';
import { int, num, numeric, optNum, scalar, spec, str } from './helpers.js';
import type { FunctionSpec } from './types.js';

const C = 'Engineering';
const BASES = { BIN: 2, OCT: 8, DEC: 10, HEX: 16 } as const;
type Base = keyof typeof BASES;
/** Two's-complement widths: 10 binary digits, 30 octal bits (10 digits), 40 hex bits (10 digits). */
const BITS: Record<Base, number> = { BIN: 10, OCT: 30, DEC: 0, HEX: 40 };

function parseBase(value: Value | undefined, base: Base): number {
	const text = toText(scalar(value)).trim().toUpperCase();
	if (base === 'DEC') return Math.trunc(num(value));
	const radix = BASES[base];
	const digits = radix === 2 ? /^[01]{1,10}$/ : radix === 8 ? /^[0-7]{1,10}$/ : /^[0-9A-F]{1,10}$/;
	if (text === '') return 0;
	if (!digits.test(text)) fail(ERR.NUM);
	let n = parseInt(text, radix);
	const bits = BITS[base];
	if (text.length === 10 && n >= Math.pow(2, bits - 1)) n -= Math.pow(2, bits);
	return n;
}

function formatBase(n: number, base: Base, places: number | undefined): string {
	const radix = BASES[base];
	const bits = BITS[base];
	const limit = Math.pow(2, bits - 1);
	if (n < -limit || n >= limit) fail(ERR.NUM);
	if (n < 0) return (Math.pow(2, bits) + n).toString(radix).toUpperCase();
	const text = n.toString(radix).toUpperCase();
	if (places === undefined) return text;
	if (places < text.length || places > 10) fail(ERR.NUM);
	return text.padStart(Math.trunc(places), '0');
}

const conversions: FunctionSpec[] = [];
for (const from of Object.keys(BASES) as Base[]) {
	for (const to of Object.keys(BASES) as Base[]) {
		if (from === to) continue;
		const name = `${from}2${to}`;
		conversions.push(
			spec(
				name,
				C,
				to === 'DEC' ? `${name}(number)` : `${name}(number, [places])`,
				`Converts a ${from.toLowerCase()} number to ${to.toLowerCase()}.`,
				1,
				to === 'DEC' ? 1 : 2,
				(args) => {
					const n = parseBase(args[0], from);
					if (to === 'DEC') return n;
					const places = args.length > 1 && args[1] !== null ? int(args[1]) : undefined;
					return formatBase(n, to, places);
				},
			),
		);
	}
}

const ROMAN: [number, string][] = [
	[1000, 'M'],
	[900, 'CM'],
	[500, 'D'],
	[400, 'CD'],
	[100, 'C'],
	[90, 'XC'],
	[50, 'L'],
	[40, 'XL'],
	[10, 'X'],
	[9, 'IX'],
	[5, 'V'],
	[4, 'IV'],
	[1, 'I'],
];

const bitArg = (value: Value | undefined): number => {
	const n = num(value);
	if (n < 0 || n >= 2 ** 48 || !Number.isInteger(n)) fail(ERR.NUM);
	return n;
};

const bitwise = (a: number, b: number, op: (x: number, y: number) => number): number => {
	const hi = op(Math.floor(a / 2 ** 24), Math.floor(b / 2 ** 24));
	const lo = op(a % 2 ** 24, b % 2 ** 24);
	return hi * 2 ** 24 + lo;
};

export const ENGINEERING_FUNCTIONS: FunctionSpec[] = [
	...conversions,
	spec('BITAND', C, 'BITAND(number1, number2)', 'Bitwise AND.', 2, 2, (args) =>
		bitwise(bitArg(args[0]), bitArg(args[1]), (x, y) => x & y),
	),
	spec('BITOR', C, 'BITOR(number1, number2)', 'Bitwise OR.', 2, 2, (args) =>
		bitwise(bitArg(args[0]), bitArg(args[1]), (x, y) => x | y),
	),
	spec('BITXOR', C, 'BITXOR(number1, number2)', 'Bitwise XOR.', 2, 2, (args) =>
		bitwise(bitArg(args[0]), bitArg(args[1]), (x, y) => x ^ y),
	),
	spec('BITLSHIFT', C, 'BITLSHIFT(number, shift_amount)', 'Shifts bits left.', 2, 2, (args) => {
		const out = Math.floor(bitArg(args[0]) * Math.pow(2, int(args[1])));
		return out >= 2 ** 48 ? fail(ERR.NUM) : out;
	}),
	spec('BITRSHIFT', C, 'BITRSHIFT(number, shift_amount)', 'Shifts bits right.', 2, 2, (args) => {
		const out = Math.floor(bitArg(args[0]) / Math.pow(2, int(args[1])));
		return out >= 2 ** 48 ? fail(ERR.NUM) : out;
	}),
	numeric(
		'DELTA',
		C,
		'DELTA(number1, [number2])',
		'1 when two numbers are equal, else 0.',
		1,
		2,
		(a, b = 0) => (a === b ? 1 : 0),
	),
	numeric(
		'GESTEP',
		C,
		'GESTEP(number, [step])',
		'1 when a number is at least step, else 0.',
		1,
		2,
		(a, b = 0) => (a >= b ? 1 : 0),
	),
	spec(
		'ROMAN',
		'Math & Trig',
		'ROMAN(number, [form])',
		'Converts a number to Roman numerals (classic form).',
		1,
		2,
		(args) => {
			let n = Math.trunc(num(args[0]));
			if (n < 0 || n > 3999) fail(ERR.VALUE);
			let out = '';
			for (const [value, letters] of ROMAN) {
				while (n >= value) {
					out += letters;
					n -= value;
				}
			}
			return out;
		},
	),
	spec(
		'ARABIC',
		'Math & Trig',
		'ARABIC(text)',
		'Converts Roman numerals to a number.',
		1,
		1,
		(args) => {
			const text = str(args[0]).trim().toUpperCase();
			const negative = text.startsWith('-');
			const body = negative ? text.slice(1) : text;
			if (!/^[MDCLXVI]*$/.test(body)) fail(ERR.VALUE);
			const value: Record<string, number> = { M: 1000, D: 500, C: 100, L: 50, X: 10, V: 5, I: 1 };
			let total = 0;
			for (let i = 0; i < body.length; i++) {
				const cur = value[body[i] ?? ''] ?? 0;
				const next = value[body[i + 1] ?? ''] ?? 0;
				total += cur < next ? -cur : cur;
			}
			return negative ? -total : total;
		},
	),
	spec(
		'BASE',
		'Math & Trig',
		'BASE(number, radix, [min_length])',
		'Converts a number to text in a radix.',
		2,
		3,
		(args) => {
			const n = Math.trunc(num(args[0]));
			const radix = Math.trunc(num(args[1]));
			if (n < 0 || n >= 2 ** 53 || radix < 2 || radix > 36) fail(ERR.NUM);
			return n
				.toString(radix)
				.toUpperCase()
				.padStart(Math.trunc(optNum(args, 2, 0)), '0');
		},
	),
	spec(
		'DECIMAL',
		'Math & Trig',
		'DECIMAL(text, radix)',
		'Converts text in a radix to a number.',
		2,
		2,
		(args) => {
			const text = str(args[0]).trim().toLowerCase();
			const radix = Math.trunc(num(args[1]));
			if (radix < 2 || radix > 36) fail(ERR.NUM);
			let total = 0;
			for (const ch of text) {
				const d = parseInt(ch, 36);
				if (Number.isNaN(d) || d >= radix) fail(ERR.NUM);
				total = total * radix + d;
			}
			return total;
		},
	),
];
