// BESSELJ, BESSELY, BESSELI and BESSELK.
// J and Y use Miller's backward recurrence and the Neumann series; I uses the periodic
// integral; K uses the trapezoid rule on K_n(x) = integral of exp(-x cosh t) cosh(n t).
import { ERR, fail } from '../values';
import { finite, num, spec } from './helpers';
import type { FunctionSpec } from './types';

const C = 'Engineering';
const EULER = 0.5772156649015329;
const LIMIT = 1e6;

/** J_0(x) .. J_m(x) for x > 0 by Miller's algorithm, normalised by J0 + 2 sum J_2k = 1. */
function besselJAll(x: number, nmax: number): number[] {
	const top = Math.max(nmax, Math.ceil(x));
	let m = top + 40 + Math.ceil(2 * Math.sqrt(top));
	if (m % 2 === 1) m += 1;
	const j = new Array<number>(m + 2).fill(0);
	j[m] = 1e-30;
	for (let k = m; k >= 1; k--) {
		j[k - 1] = ((2 * k) / x) * (j[k] as number) - (j[k + 1] as number);
		if (Math.abs(j[k - 1] as number) > 1e250) {
			for (let i = k - 1; i <= m + 1; i++) j[i] = (j[i] as number) * 1e-250;
		}
	}
	let sum = j[0] as number;
	for (let k = 2; k <= m; k += 2) sum += 2 * (j[k] as number);
	return j.slice(0, m + 1).map((v) => v / sum);
}

export function besselJ(x: number, n: number): number {
	if (x === 0) return n === 0 ? 1 : 0;
	const value = besselJAll(Math.abs(x), n)[n] ?? 0;
	return x < 0 && n % 2 === 1 ? -value : value;
}

export function besselY(x: number, n: number): number {
	const j = besselJAll(x, n + 2);
	const j0 = j[0] as number;
	const j1 = j[1] as number;
	const lg = Math.log(x / 2) + EULER;
	let s0 = 0;
	let s1 = 0;
	for (let k = 1; 2 * k + 1 < j.length; k++) {
		const sign = k % 2 === 1 ? 1 : -1;
		s0 += (sign * (j[2 * k] as number)) / k;
		s1 += (sign * ((j[2 * k - 1] as number) - (j[2 * k + 1] as number))) / k;
	}
	let prev = (2 / Math.PI) * lg * j0 + (4 / Math.PI) * s0;
	if (n === 0) return prev;
	let cur = (2 / Math.PI) * (lg * j1 - j0 / x) - (2 / Math.PI) * s1;
	for (let k = 1; k < n; k++) {
		const next = ((2 * k) / x) * cur - prev;
		prev = cur;
		cur = next;
	}
	return cur;
}

export function besselI(x: number, n: number): number {
	const ax = Math.abs(x);
	if (ax > 700) fail(ERR.NUM);
	if (ax === 0) return n === 0 ? 1 : 0;
	const points = 2 * (Math.ceil(ax) + n) + 80;
	let sum = 0;
	for (let i = 0; i < points; i++) {
		const t = (2 * Math.PI * i) / points;
		sum += Math.exp(ax * (Math.cos(t) - 1)) * Math.cos(n * t);
	}
	const value = (sum / points) * Math.exp(ax);
	return x < 0 && n % 2 === 1 ? -value : value;
}

export function besselK(x: number, n: number): number {
	const h = 0.1;
	let sum = 0.5 * Math.exp(-x);
	for (let i = 1; i < 100000; i++) {
		const t = i * h;
		const term = Math.exp(-x * Math.cosh(t)) * Math.cosh(n * t);
		sum += term;
		if (term < sum * 1e-18 && t > n) break;
	}
	return sum * h;
}

const bessel = (
	name: string,
	description: string,
	positiveOnly: boolean,
	fn: (x: number, n: number) => number,
): FunctionSpec =>
	spec(name, C, `${name}(x, n)`, description, 2, 2, (args) => {
		const x = num(args[0]);
		const n = Math.trunc(num(args[1]));
		if (n < 0 || (positiveOnly && x <= 0) || Math.abs(x) > LIMIT) fail(ERR.NUM);
		return finite(fn(x, n));
	});

export const BESSEL_FUNCTIONS: FunctionSpec[] = [
	bessel('BESSELJ', 'The Bessel function Jn(x).', false, besselJ),
	bessel('BESSELY', 'The Bessel function Yn(x), also called the Weber function.', true, besselY),
	bessel('BESSELI', 'The modified Bessel function In(x).', false, besselI),
	bessel('BESSELK', 'The modified Bessel function Kn(x).', true, besselK),
];
