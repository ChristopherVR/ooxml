import { round15 } from '../text-number.js';
import { ERR, fail } from '../values.js';
import { numeric } from './helpers.js';
import { domain } from './helpers.js';
import { SUM_FUNCTIONS } from './math-sum.js';
import type { FunctionSpec } from './types.js';
import { MATH_MORE } from './math-more.js';
import { combin, COMBINATORIC_FUNCTIONS } from './math-combin.js';
import { excelMod, mround } from './math-mod.js';
import { power } from '../operators.js';

export { combin };

const C = 'Math & Trig';

/** Rounds `x` to `digits` decimals (negative: to tens, hundreds) at Excel's 15-digit precision. */
export function roundTo(x: number, digits: number, mode: 'half' | 'up' | 'down'): number {
	if (!Number.isFinite(x)) return x;
	const d = Math.trunc(digits);
	if (d > 15) return x;
	const sign = x < 0 ? -1 : 1;
	const factor = Math.pow(10, Math.abs(d));
	const scaled = round15(d >= 0 ? Math.abs(x) * factor : Math.abs(x) / factor);
	const rounded =
		mode === 'half'
			? Math.floor(scaled + 0.5)
			: mode === 'up'
				? Math.ceil(scaled)
				: Math.floor(scaled);
	const result = d >= 0 ? rounded / factor : rounded * factor;
	return result === 0 ? 0 : sign * result;
}

const quotientMultiple = (x: number, sig: number, op: (v: number) => number): number => {
	if (sig === 0) return 0;
	return op(round15(x / sig)) * sig;
};

export const MATH_FUNCTIONS: FunctionSpec[] = [
	...MATH_MORE,
	...SUM_FUNCTIONS,
	numeric('ABS', C, 'ABS(number)', 'The absolute value of a number.', 1, 1, Math.abs),
	numeric('SIGN', C, 'SIGN(number)', 'The sign of a number: 1, 0 or -1.', 1, 1, (x) =>
		Math.sign(x),
	),
	numeric(
		'INT',
		C,
		'INT(number)',
		'Rounds a number down to the nearest integer.',
		1,
		1,
		Math.floor,
	),
	numeric(
		'TRUNC',
		C,
		'TRUNC(number, [num_digits])',
		'Truncates a number toward zero.',
		1,
		2,
		(x, d) => roundTo(x, d ?? 0, 'down'),
	),
	numeric(
		'ROUND',
		C,
		'ROUND(number, num_digits)',
		'Rounds a number to a number of digits.',
		2,
		2,
		(x, d) => roundTo(x, d ?? 0, 'half'),
	),
	numeric(
		'ROUNDUP',
		C,
		'ROUNDUP(number, num_digits)',
		'Rounds a number away from zero.',
		2,
		2,
		(x, d) => roundTo(x, d ?? 0, 'up'),
	),
	numeric(
		'ROUNDDOWN',
		C,
		'ROUNDDOWN(number, num_digits)',
		'Rounds a number toward zero.',
		2,
		2,
		(x, d) => roundTo(x, d ?? 0, 'down'),
	),
	numeric(
		'MROUND',
		C,
		'MROUND(number, multiple)',
		'Rounds a number to the nearest multiple.',
		2,
		2,
		(x, m) => mround(x, m ?? 0),
	),
	numeric(
		'CEILING',
		C,
		'CEILING(number, significance)',
		'Rounds a number up to a multiple of significance.',
		2,
		2,
		(x, s) => {
			if (x > 0 && (s ?? 0) < 0) fail(ERR.NUM);
			return quotientMultiple(x, s ?? 0, Math.ceil);
		},
	),
	numeric(
		'FLOOR',
		C,
		'FLOOR(number, significance)',
		'Rounds a number down to a multiple of significance.',
		2,
		2,
		(x, s) => {
			if (s === 0) fail(ERR.DIV0);
			if (x > 0 && (s ?? 0) < 0) fail(ERR.NUM);
			return quotientMultiple(x, s ?? 1, Math.floor);
		},
	),
	numeric(
		'CEILING.MATH',
		C,
		'CEILING.MATH(number, [significance], [mode])',
		'Rounds a number up to an integer or multiple.',
		1,
		3,
		(x, s, mode) => {
			const sig = Math.abs(s ?? 1);
			if (x < 0 && mode) return -quotientMultiple(-x, sig, Math.ceil);
			return quotientMultiple(x, sig, Math.ceil);
		},
	),
	numeric(
		'FLOOR.MATH',
		C,
		'FLOOR.MATH(number, [significance], [mode])',
		'Rounds a number down to an integer or multiple.',
		1,
		3,
		(x, s, mode) => {
			const sig = Math.abs(s ?? 1);
			if (x < 0 && mode) return -quotientMultiple(-x, sig, Math.floor);
			return quotientMultiple(x, sig, Math.floor);
		},
	),
	numeric(
		'CEILING.PRECISE',
		C,
		'CEILING.PRECISE(number, [significance])',
		'Rounds a number up regardless of sign.',
		1,
		2,
		(x, s) => quotientMultiple(x, Math.abs(s ?? 1), Math.ceil),
	),
	numeric(
		'ISO.CEILING',
		C,
		'ISO.CEILING(number, [significance])',
		'Rounds a number up regardless of sign.',
		1,
		2,
		(x, s) => quotientMultiple(x, Math.abs(s ?? 1), Math.ceil),
	),
	numeric(
		'FLOOR.PRECISE',
		C,
		'FLOOR.PRECISE(number, [significance])',
		'Rounds a number down regardless of sign.',
		1,
		2,
		(x, s) => quotientMultiple(x, Math.abs(s ?? 1), Math.floor),
	),
	numeric(
		'EVEN',
		C,
		'EVEN(number)',
		'Rounds a number away from zero to an even integer.',
		1,
		1,
		(x) => Math.sign(x) * Math.ceil(round15(Math.abs(x)) / 2) * 2,
	),
	numeric(
		'ODD',
		C,
		'ODD(number)',
		'Rounds a number away from zero to an odd integer.',
		1,
		1,
		(x) => {
			const n = Math.ceil(round15(Math.abs(x)));
			return (x < 0 ? -1 : 1) * (n % 2 === 0 ? n + 1 : n);
		},
	),
	numeric(
		'MOD',
		C,
		'MOD(number, divisor)',
		'The remainder after division (sign of the divisor).',
		2,
		2,
		(n, d) => excelMod(n, d ?? 0),
	),
	numeric(
		'QUOTIENT',
		C,
		'QUOTIENT(numerator, denominator)',
		'The integer part of a division.',
		2,
		2,
		(n, d) => {
			if (d === 0) fail(ERR.DIV0);
			return Math.trunc(n / (d ?? 1));
		},
	),
	numeric('SQRT', C, 'SQRT(number)', 'The positive square root.', 1, 1, (x) =>
		domain(x >= 0, Math.sqrt(x)),
	),
	numeric('SQRTPI', C, 'SQRTPI(number)', 'The square root of number * pi.', 1, 1, (x) =>
		domain(x >= 0, Math.sqrt(x * Math.PI)),
	),
	numeric('POWER', C, 'POWER(number, power)', 'A number raised to a power.', 2, 2, (b, e) =>
		power(b, e ?? 1),
	),
	numeric('EXP', C, 'EXP(number)', 'e raised to a power.', 1, 1, Math.exp),
	numeric('LN', C, 'LN(number)', 'The natural logarithm.', 1, 1, (x) => domain(x > 0, Math.log(x))),
	numeric('LOG10', C, 'LOG10(number)', 'The base-10 logarithm.', 1, 1, (x) =>
		domain(x > 0, Math.log10(x)),
	),
	numeric(
		'LOG',
		C,
		'LOG(number, [base])',
		'The logarithm to a base (10 by default).',
		1,
		2,
		(x, b) => {
			const base = b ?? 10;
			if (x <= 0 || base <= 0) fail(ERR.NUM);
			if (base === 1) fail(ERR.DIV0);
			return base === 10 ? Math.log10(x) : Math.log(x) / Math.log(base);
		},
	),
	numeric('PI', C, 'PI()', 'The value of pi.', 0, 0, () => Math.PI),
	...COMBINATORIC_FUNCTIONS,
];
