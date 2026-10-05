import { ERR, fail } from '../values.js';
import { collectNumbers, num, numeric, spec } from './helpers.js';
import { domain } from './helpers.js';
import type { FunctionSpec } from './types.js';

const C = 'Math & Trig';

/** Trigonometry, random numbers and series. */
export const MATH_MORE: FunctionSpec[] = [
	numeric('SIN', C, 'SIN(number)', 'The sine of an angle in radians.', 1, 1, Math.sin),
	numeric('COS', C, 'COS(number)', 'The cosine of an angle in radians.', 1, 1, Math.cos),
	numeric('TAN', C, 'TAN(number)', 'The tangent of an angle in radians.', 1, 1, Math.tan),
	numeric('ASIN', C, 'ASIN(number)', 'The arcsine in radians.', 1, 1, (x) =>
		domain(Math.abs(x) <= 1, Math.asin(x)),
	),
	numeric('ACOS', C, 'ACOS(number)', 'The arccosine in radians.', 1, 1, (x) =>
		domain(Math.abs(x) <= 1, Math.acos(x)),
	),
	numeric('ATAN', C, 'ATAN(number)', 'The arctangent in radians.', 1, 1, Math.atan),
	numeric(
		'ATAN2',
		C,
		'ATAN2(x_num, y_num)',
		'The arctangent of x and y coordinates.',
		2,
		2,
		(x, y) => {
			if (x === 0 && y === 0) fail(ERR.DIV0);
			return Math.atan2(y ?? 0, x);
		},
	),
	numeric('SINH', C, 'SINH(number)', 'The hyperbolic sine.', 1, 1, Math.sinh),
	numeric('COSH', C, 'COSH(number)', 'The hyperbolic cosine.', 1, 1, Math.cosh),
	numeric('TANH', C, 'TANH(number)', 'The hyperbolic tangent.', 1, 1, Math.tanh),
	numeric('ASINH', C, 'ASINH(number)', 'The inverse hyperbolic sine.', 1, 1, Math.asinh),
	numeric('ACOSH', C, 'ACOSH(number)', 'The inverse hyperbolic cosine.', 1, 1, (x) =>
		domain(x >= 1, Math.acosh(x)),
	),
	numeric('ATANH', C, 'ATANH(number)', 'The inverse hyperbolic tangent.', 1, 1, (x) =>
		domain(Math.abs(x) < 1, Math.atanh(x)),
	),
	numeric('COT', C, 'COT(number)', 'The cotangent of an angle.', 1, 1, (x) =>
		x === 0 ? fail(ERR.DIV0) : 1 / Math.tan(x),
	),
	numeric('COTH', C, 'COTH(number)', 'The hyperbolic cotangent.', 1, 1, (x) =>
		x === 0 ? fail(ERR.DIV0) : 1 / Math.tanh(x),
	),
	numeric('CSC', C, 'CSC(number)', 'The cosecant of an angle.', 1, 1, (x) =>
		x === 0 ? fail(ERR.DIV0) : 1 / Math.sin(x),
	),
	numeric('CSCH', C, 'CSCH(number)', 'The hyperbolic cosecant.', 1, 1, (x) =>
		x === 0 ? fail(ERR.DIV0) : 1 / Math.sinh(x),
	),
	numeric('SEC', C, 'SEC(number)', 'The secant of an angle.', 1, 1, (x) => 1 / Math.cos(x)),
	numeric('SECH', C, 'SECH(number)', 'The hyperbolic secant.', 1, 1, (x) => 1 / Math.cosh(x)),
	numeric('ACOT', C, 'ACOT(number)', 'The arccotangent.', 1, 1, (x) => Math.PI / 2 - Math.atan(x)),
	numeric('ACOTH', C, 'ACOTH(number)', 'The inverse hyperbolic cotangent.', 1, 1, (x) =>
		domain(Math.abs(x) > 1, 0.5 * Math.log((x + 1) / (x - 1))),
	),
	numeric(
		'DEGREES',
		C,
		'DEGREES(angle)',
		'Converts radians to degrees.',
		1,
		1,
		(x) => (x * 180) / Math.PI,
	),
	numeric(
		'RADIANS',
		C,
		'RADIANS(angle)',
		'Converts degrees to radians.',
		1,
		1,
		(x) => (x * Math.PI) / 180,
	),
	spec(
		'RAND',
		C,
		'RAND()',
		'A random number between 0 and 1.',
		0,
		0,
		(_args, ctx) => ctx.frame.host.random(),
		undefined,
		true,
	),
	spec(
		'RANDBETWEEN',
		C,
		'RANDBETWEEN(bottom, top)',
		'A random integer between two numbers.',
		2,
		2,
		(args, ctx) => {
			const lo = Math.ceil(num(args[0]));
			const hi = Math.floor(num(args[1]));
			if (lo > hi) fail(ERR.NUM);
			return lo + Math.floor(ctx.frame.host.random() * (hi - lo + 1));
		},
		undefined,
		true,
	),
	spec(
		'SERIESSUM',
		C,
		'SERIESSUM(x, n, m, coefficients)',
		'The sum of a power series.',
		4,
		4,
		(args, ctx) => {
			const x = num(args[0]);
			const n = num(args[1]);
			const m = num(args[2]);
			return collectNumbers(ctx, [args[3] ?? null]).reduce(
				(acc, a, i) => acc + a * Math.pow(x, n + i * m),
				0,
			);
		},
		['value', 'value', 'value', 'any'],
	),
];
