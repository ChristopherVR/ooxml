// MOD and MROUND with Excel's precision and limits.
import { ERR, fail } from '../values.js';

const SPLIT = 134217729; // 2^27 + 1

/** The rounding error of `a * b` (Dekker's two-product), so `a * b` is exactly `p + error`. */
function productError(a: number, b: number, p: number): number {
	const ca = SPLIT * a;
	const ah = ca - (ca - a);
	const al = a - ah;
	const cb = SPLIT * b;
	const bh = cb - (cb - b);
	const bl = b - bh;
	return ah * bh - p + ah * bl + al * bh + al * bl;
}

/** `n - q * d` computed from the exact product, rounded once. */
function remainder(n: number, d: number, q: number): number {
	const p = q * d;
	return n - p - productError(q, d, p);
}

/** Excel refuses MOD when the quotient reaches this size (about 2^50 / 1000). */
const MOD_LIMIT = 1.1259e12;

/**
 * MOD like Excel: the remainder takes the divisor's sign and is computed from the exact quotient
 * (MOD(0.3, 0.1) is 0.09999999999999998, not 0), and huge quotients are `#NUM!`.
 */
export function excelMod(n: number, d: number): number {
	if (d === 0) fail(ERR.DIV0);
	const ratio = n / d;
	if (!(Math.abs(ratio) < MOD_LIMIT)) fail(ERR.NUM);
	let q = Math.floor(ratio);
	let r = remainder(n, d, q);
	if (r !== 0 && r < 0 !== d < 0) r = remainder(n, d, --q);
	else if (Math.abs(r) >= Math.abs(d)) r = remainder(n, d, ++q);
	return r === 0 ? 0 : r;
}

/** Excel rounds a quotient up when it is within 2^-48 below the half (MROUND(6.05, 0.1) is 6). */
const HALF = 0.5 - 2 ** -48;

/** MROUND: the nearest multiple of `m` (halves away from zero); signs must agree. */
export function mround(x: number, m: number): number {
	if (m === 0 || x === 0) return 0;
	if (x < 0 !== m < 0) fail(ERR.NUM);
	const q = x / m;
	const whole = Math.floor(q);
	const n = q - whole >= HALF ? whole + 1 : whole;
	const result = n * m;
	return result === 0 ? 0 : result;
}
