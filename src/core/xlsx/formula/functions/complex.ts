import { toText } from '../coerce';
import type { CallContext } from '../context';
import { ERR, fail, isError, type Scalar, type Value } from '../values';
import { num, scalar, spec } from './helpers';
import type { FunctionSpec } from './types';

const C = 'Engineering';

interface Complex {
	re: number;
	im: number;
	suffix: 'i' | 'j';
}

const NUMBER = String.raw`(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?`;
const REAL_ONLY = new RegExp(`^([+-]?${NUMBER})$`);
const IMAGINARY_ONLY = new RegExp(`^([+-]?(?:${NUMBER})?)([ij])$`);
const BOTH = new RegExp(`^([+-]?${NUMBER})([+-])((?:${NUMBER})?)([ij])$`);

const coefficient = (text: string): number =>
	text === '' || text === '+' ? 1 : text === '-' ? -1 : Number(text);

/** A complex number from Excel text such as `3+4i`, `-j` or `5`; anything else is `#NUM!`. */
export function parseComplex(value: Scalar): Complex {
	if (isError(value)) fail(value);
	if (typeof value === 'number') return { re: value, im: 0, suffix: 'i' };
	const text = toText(value);
	if (text === '') return { re: 0, im: 0, suffix: 'i' };
	const real = REAL_ONLY.exec(text);
	if (real) return { re: Number(real[1]), im: 0, suffix: 'i' };
	const imaginary = IMAGINARY_ONLY.exec(text);
	if (imaginary)
		return {
			re: 0,
			im: coefficient(imaginary[1] ?? ''),
			suffix: imaginary[2] === 'j' ? 'j' : 'i',
		};
	const both = BOTH.exec(text);
	if (!both) return fail(ERR.NUM);
	const sign = both[2] === '-' ? -1 : 1;
	return {
		re: Number(both[1]),
		im: sign * coefficient(both[3] ?? ''),
		suffix: both[4] === 'j' ? 'j' : 'i',
	};
}

/** 15 significant digits, with Excel's `E+05` exponent form. */
function digits(n: number): string {
	const text = Number(n.toPrecision(15)).toString();
	const match = /^(-?[\d.]+)e([+-])(\d+)$/.exec(text);
	return match ? `${match[1]}E${match[2]}${(match[3] ?? '').padStart(2, '0')}` : text;
}

export function formatComplex({ re, im, suffix }: Complex): string {
	const r = Math.abs(re) < 1e-15 * Math.max(1, Math.abs(im)) ? 0 : re;
	const i = Math.abs(im) < 1e-15 * Math.max(1, Math.abs(re)) ? 0 : im;
	if (i === 0) return digits(r);
	const imaginary = i === 1 ? suffix : i === -1 ? `-${suffix}` : `${digits(i)}${suffix}`;
	if (r === 0) return imaginary;
	return `${digits(r)}${i > 0 ? '+' : ''}${imaginary}`;
}

const make = (re: number, im: number, suffix: 'i' | 'j'): Complex => ({ re, im, suffix });
const modulus = (z: Complex): number => Math.hypot(z.re, z.im);
const argument = (z: Complex): number => Math.atan2(z.im, z.re);
const polar = (r: number, theta: number, s: 'i' | 'j'): Complex =>
	make(r * Math.cos(theta), r * Math.sin(theta), s);

function divide(a: Complex, b: Complex): Complex {
	const d = b.re * b.re + b.im * b.im;
	if (d === 0) fail(ERR.NUM);
	return make((a.re * b.re + a.im * b.im) / d, (a.im * b.re - a.re * b.im) / d, a.suffix);
}

/** `z^n` by repeated multiplication, which keeps small integer powers exact. */
function integerPower(z: Complex, n: number): Complex {
	let result = one(z);
	for (let i = 0; i < Math.abs(n); i++)
		result = make(
			result.re * z.re - result.im * z.im,
			result.re * z.im + result.im * z.re,
			z.suffix,
		);
	return n < 0 ? divide(one(z), result) : result;
}

const exp = (z: Complex): Complex => polar(Math.exp(z.re), z.im, z.suffix);
function ln(z: Complex): Complex {
	if (z.re === 0 && z.im === 0) fail(ERR.NUM);
	return make(Math.log(modulus(z)), argument(z), z.suffix);
}
const scale = (z: Complex, k: number): Complex => make(z.re * k, z.im * k, z.suffix);
const sin = (z: Complex): Complex =>
	make(Math.sin(z.re) * Math.cosh(z.im), Math.cos(z.re) * Math.sinh(z.im), z.suffix);
const cos = (z: Complex): Complex =>
	make(Math.cos(z.re) * Math.cosh(z.im), -Math.sin(z.re) * Math.sinh(z.im), z.suffix);
const sinh = (z: Complex): Complex =>
	make(Math.sinh(z.re) * Math.cos(z.im), Math.cosh(z.re) * Math.sin(z.im), z.suffix);
const cosh = (z: Complex): Complex =>
	make(Math.cosh(z.re) * Math.cos(z.im), Math.sinh(z.re) * Math.sin(z.im), z.suffix);
const one = (z: Complex): Complex => make(1, 0, z.suffix);
const reciprocal = (z: Complex): Complex => divide(one(z), z);

/** Both operands must use the same suffix, or the result is `#VALUE!`. */
function joined(a: Complex, b: Complex): 'i' | 'j' {
	if (a.suffix !== b.suffix && a.im !== 0 && b.im !== 0) fail(ERR.VALUE);
	return a.im !== 0 ? a.suffix : b.suffix;
}

const unary = (
	name: string,
	description: string,
	fn: (z: Complex) => Complex | number,
): FunctionSpec =>
	spec(name, C, `${name}(inumber)`, description, 1, 1, (args) => {
		const result = fn(parseComplex(scalar(args[0])));
		return typeof result === 'number' ? result : formatComplex(result);
	});

const values = (ctx: CallContext, args: Value[]): Complex[] => {
	const out: Complex[] = [];
	for (const arg of args) ctx.forEach(arg, (v) => out.push(parseComplex(v)));
	return out;
};

const fold = (
	name: string,
	description: string,
	identity: Complex,
	step: (acc: Complex, next: Complex) => Complex,
): FunctionSpec =>
	spec(
		name,
		C,
		`${name}(inumber1, [inumber2], ...)`,
		description,
		1,
		255,
		(args, ctx) => {
			const list = values(ctx, args);
			let acc = list[0] ?? identity;
			for (const next of list.slice(1)) acc = step(acc, next);
			return formatComplex(acc);
		},
		['any'],
	);

/** Complex numbers held as text (`COMPLEX` and the `IM` functions). */
export const COMPLEX_FUNCTIONS: FunctionSpec[] = [
	spec(
		'COMPLEX',
		C,
		'COMPLEX(real_num, i_num, [suffix])',
		'Builds a complex number from its real and imaginary coefficients.',
		2,
		3,
		(args) => {
			const suffix = args.length > 2 && args[2] !== null ? toText(scalar(args[2])) : 'i';
			if (suffix !== 'i' && suffix !== 'j') fail(ERR.VALUE);
			return formatComplex(make(num(args[0]), num(args[1]), suffix));
		},
	),
	unary('IMREAL', 'The real coefficient of a complex number.', (z) => z.re),
	unary('IMAGINARY', 'The imaginary coefficient of a complex number.', (z) => z.im),
	unary('IMABS', 'The absolute value (modulus) of a complex number.', modulus),
	unary('IMARGUMENT', 'The argument theta, an angle in radians.', (z) =>
		z.re === 0 && z.im === 0 ? fail(ERR.DIV0) : argument(z),
	),
	unary('IMCONJUGATE', 'The complex conjugate.', (z) => make(z.re, -z.im, z.suffix)),
	unary('IMEXP', 'The exponential of a complex number.', exp),
	unary('IMLN', 'The natural logarithm of a complex number.', ln),
	unary('IMLOG10', 'The base-10 logarithm of a complex number.', (z) =>
		scale(ln(z), 1 / Math.LN10),
	),
	unary('IMLOG2', 'The base-2 logarithm of a complex number.', (z) => scale(ln(z), 1 / Math.LN2)),
	unary('IMSQRT', 'The square root of a complex number.', (z) =>
		polar(Math.sqrt(modulus(z)), argument(z) / 2, z.suffix),
	),
	unary('IMSIN', 'The sine of a complex number.', sin),
	unary('IMCOS', 'The cosine of a complex number.', cos),
	unary('IMTAN', 'The tangent of a complex number.', (z) => divide(sin(z), cos(z))),
	unary('IMSEC', 'The secant of a complex number.', (z) => reciprocal(cos(z))),
	unary('IMCSC', 'The cosecant of a complex number.', (z) => reciprocal(sin(z))),
	unary('IMCOT', 'The cotangent of a complex number.', (z) => divide(cos(z), sin(z))),
	unary('IMSINH', 'The hyperbolic sine of a complex number.', sinh),
	unary('IMCOSH', 'The hyperbolic cosine of a complex number.', cosh),
	unary('IMSECH', 'The hyperbolic secant of a complex number.', (z) => reciprocal(cosh(z))),
	unary('IMCSCH', 'The hyperbolic cosecant of a complex number.', (z) => reciprocal(sinh(z))),
	spec(
		'IMPOWER',
		C,
		'IMPOWER(inumber, number)',
		'A complex number raised to a power.',
		2,
		2,
		(args) => {
			const z = parseComplex(scalar(args[0]));
			const n = num(args[1]);
			if (z.re === 0 && z.im === 0) return n > 0 ? '0' : fail(ERR.NUM);
			if (Number.isInteger(n) && Math.abs(n) <= 64) return formatComplex(integerPower(z, n));
			return formatComplex(polar(modulus(z) ** n, argument(z) * n, z.suffix));
		},
	),
	spec(
		'IMDIV',
		C,
		'IMDIV(inumber1, inumber2)',
		'The quotient of two complex numbers.',
		2,
		2,
		(args) => {
			const a = parseComplex(scalar(args[0]));
			const b = parseComplex(scalar(args[1]));
			return formatComplex({ ...divide(a, b), suffix: joined(a, b) });
		},
	),
	fold('IMSUM', 'The sum of complex numbers.', make(0, 0, 'i'), (a, b) =>
		make(a.re + b.re, a.im + b.im, joined(a, b)),
	),
	fold('IMPRODUCT', 'The product of complex numbers.', make(1, 0, 'i'), (a, b) =>
		make(a.re * b.re - a.im * b.im, a.re * b.im + a.im * b.re, joined(a, b)),
	),
	spec(
		'IMSUB',
		C,
		'IMSUB(inumber1, inumber2)',
		'The difference of two complex numbers.',
		2,
		2,
		(args) => {
			const a = parseComplex(scalar(args[0]));
			const b = parseComplex(scalar(args[1]));
			return formatComplex(make(a.re - b.re, a.im - b.im, joined(a, b)));
		},
	),
];
