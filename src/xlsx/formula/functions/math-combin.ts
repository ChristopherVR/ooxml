// Factorials, combinations, GCD and LCM.
import { ERR, fail, type Value } from '../values.js';
import { collectNumbers, numeric, spec } from './helpers.js';
import type { FunctionSpec } from './types.js';

const C = 'Math & Trig';

export const fact = (n: number): number => {
	if (n < 0) fail(ERR.NUM);
	let out = 1;
	for (let i = 2; i <= Math.trunc(n); i++) out *= i;
	return out;
};

export function combin(n: number, k: number): number {
	const nn = Math.trunc(n);
	const kk = Math.trunc(k);
	if (nn < 0 || kk < 0 || kk > nn) fail(ERR.NUM);
	let out = 1;
	for (let i = 1; i <= Math.min(kk, nn - kk); i++)
		out = (out * (nn - Math.min(kk, nn - kk) + i)) / i;
	return Math.round(out);
}

const gcd2 = (a: number, b: number): number => (b === 0 ? a : gcd2(b, a % b));

function integers(args: Value[], ctx: Parameters<NonNullable<FunctionSpec['fn']>>[1]): number[] {
	return collectNumbers(ctx, args).map((n) => {
		if (n < 0) fail(ERR.NUM);
		return Math.trunc(n);
	});
}

export const COMBINATORIC_FUNCTIONS: FunctionSpec[] = [
	numeric('FACT', C, 'FACT(number)', 'The factorial of a number.', 1, 1, fact),
	numeric('FACTDOUBLE', C, 'FACTDOUBLE(number)', 'The double factorial of a number.', 1, 1, (n) => {
		if (n < -1) fail(ERR.NUM);
		let out = 1;
		for (let i = Math.trunc(n); i > 1; i -= 2) out *= i;
		return out;
	}),
	numeric(
		'COMBIN',
		C,
		'COMBIN(number, number_chosen)',
		'The number of combinations.',
		2,
		2,
		(n, k) => combin(n, k ?? 0),
	),
	numeric(
		'COMBINA',
		C,
		'COMBINA(number, number_chosen)',
		'Combinations with repetitions.',
		2,
		2,
		(n, k) => {
			if (n < 0 || (k ?? 0) < 0) fail(ERR.NUM);
			return Math.trunc(n) === 0 && Math.trunc(k ?? 0) === 0 ? 1 : combin(n + (k ?? 0) - 1, k ?? 0);
		},
	),
	spec(
		'MULTINOMIAL',
		C,
		'MULTINOMIAL(number1, ...)',
		'The multinomial of a set of numbers.',
		1,
		255,
		(args, ctx) => {
			const values = integers(args, ctx);
			const total = values.reduce((a, b) => a + b, 0);
			return values.reduce((acc, v) => acc / fact(v), fact(total));
		},
		['any'],
	),
	spec(
		'GCD',
		C,
		'GCD(number1, ...)',
		'The greatest common divisor.',
		1,
		255,
		(args, ctx) => integers(args, ctx).reduce((a, b) => gcd2(a, b), 0),
		['any'],
	),
	spec(
		'LCM',
		C,
		'LCM(number1, ...)',
		'The least common multiple.',
		1,
		255,
		(args, ctx) =>
			integers(args, ctx).reduce((a, b) => (a === 0 || b === 0 ? 0 : (a / gcd2(a, b)) * b), 1),
		['any'],
	),
];
