// Numerical core of the statistical distributions (incomplete gamma and beta, inverses).
import { ERR, fail } from '../values.js';
import { combin } from './math.js';
import { gammaLn } from './stats-core.js';

const EPS = 1e-15;

/** Regularized lower incomplete gamma P(a, x). */
export function gammaP(a: number, x: number): number {
	if (x <= 0) return 0;
	const lnPrefix = -x + a * Math.log(x) - gammaLn(a);
	if (x < a + 1) {
		let ap = a;
		let sum = 1 / a;
		let del = sum;
		for (let n = 0; n < 1000; n++) {
			ap += 1;
			del *= x / ap;
			sum += del;
			if (Math.abs(del) < Math.abs(sum) * EPS) break;
		}
		return sum * Math.exp(lnPrefix);
	}
	let b = x + 1 - a;
	let c = 1 / 1e-300;
	let d = 1 / b;
	let h = d;
	for (let i = 1; i < 1000; i++) {
		const an = -i * (i - a);
		b += 2;
		d = an * d + b;
		if (Math.abs(d) < 1e-300) d = 1e-300;
		c = b + an / c;
		if (Math.abs(c) < 1e-300) c = 1e-300;
		d = 1 / d;
		const del = d * c;
		h *= del;
		if (Math.abs(del - 1) < EPS) break;
	}
	return 1 - Math.exp(lnPrefix) * h;
}

export function betaCf(a: number, b: number, x: number): number {
	const qab = a + b;
	const qap = a + 1;
	const qam = a - 1;
	let c = 1;
	let d = 1 - (qab * x) / qap;
	if (Math.abs(d) < 1e-300) d = 1e-300;
	d = 1 / d;
	let h = d;
	for (let m = 1; m <= 1000; m++) {
		const m2 = 2 * m;
		let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
		d = 1 + aa * d;
		if (Math.abs(d) < 1e-300) d = 1e-300;
		c = 1 + aa / c;
		if (Math.abs(c) < 1e-300) c = 1e-300;
		d = 1 / d;
		h *= d * c;
		aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
		d = 1 + aa * d;
		if (Math.abs(d) < 1e-300) d = 1e-300;
		c = 1 + aa / c;
		if (Math.abs(c) < 1e-300) c = 1e-300;
		d = 1 / d;
		const del = d * c;
		h *= del;
		if (Math.abs(del - 1) < EPS) break;
	}
	return h;
}

/** Regularized incomplete beta I_x(a, b). */
export function betaI(x: number, a: number, b: number): number {
	if (x <= 0) return 0;
	if (x >= 1) return 1;
	const bt = Math.exp(
		gammaLn(a + b) - gammaLn(a) - gammaLn(b) + a * Math.log(x) + b * Math.log(1 - x),
	);
	return x < (a + 1) / (a + b + 2)
		? (bt * betaCf(a, b, x)) / a
		: 1 - (bt * betaCf(b, a, 1 - x)) / b;
}

export const erf = (x: number): number => (x < 0 ? -gammaP(0.5, x * x) : gammaP(0.5, x * x));
export const normCdf = (z: number): number => 0.5 * (1 + erf(z / Math.SQRT2));
export const normPdf = (z: number): number => Math.exp((-z * z) / 2) / Math.sqrt(2 * Math.PI);

/** Inverse standard normal CDF (Acklam's approximation refined by one Halley step). */
export function normInv(p: number): number {
	if (p <= 0 || p >= 1) fail(ERR.NUM);
	const a = [
		-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716,
		2.506628277459239,
	];
	const b = [
		-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972,
		-13.28068155288572,
	];
	const c = [
		-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734,
		4.374664141464968, 2.938163982698783,
	];
	const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
	const at = (arr: number[], i: number): number => arr[i] as number;
	let x: number;
	if (p < 0.02425) {
		const q = Math.sqrt(-2 * Math.log(p));
		x =
			(((((at(c, 0) * q + at(c, 1)) * q + at(c, 2)) * q + at(c, 3)) * q + at(c, 4)) * q +
				at(c, 5)) /
			((((at(d, 0) * q + at(d, 1)) * q + at(d, 2)) * q + at(d, 3)) * q + 1);
	} else if (p <= 1 - 0.02425) {
		const q = p - 0.5;
		const r = q * q;
		x =
			((((((at(a, 0) * r + at(a, 1)) * r + at(a, 2)) * r + at(a, 3)) * r + at(a, 4)) * r +
				at(a, 5)) *
				q) /
			(((((at(b, 0) * r + at(b, 1)) * r + at(b, 2)) * r + at(b, 3)) * r + at(b, 4)) * r + 1);
	} else {
		const q = Math.sqrt(-2 * Math.log(1 - p));
		x =
			-(
				((((at(c, 0) * q + at(c, 1)) * q + at(c, 2)) * q + at(c, 3)) * q + at(c, 4)) * q +
				at(c, 5)
			) /
			((((at(d, 0) * q + at(d, 1)) * q + at(d, 2)) * q + at(d, 3)) * q + 1);
	}
	const e = normCdf(x) - p;
	const u = e * Math.sqrt(2 * Math.PI) * Math.exp((x * x) / 2);
	return x - u / (1 + (x * u) / 2);
}

/** Inverts a monotonic increasing CDF on (lo, hi) by bisection. */
export function invert(cdf: (x: number) => number, p: number, lo: number, hi: number): number {
	let a = lo;
	let b = hi;
	while (cdf(b) < p && b < 1e10) b *= 2;
	for (let i = 0; i < 200; i++) {
		const mid = (a + b) / 2;
		if (cdf(mid) < p) a = mid;
		else b = mid;
		if (b - a < 1e-15 * Math.max(1, Math.abs(mid))) break;
	}
	return (a + b) / 2;
}

export const tCdf = (t: number, df: number): number => {
	const x = df / (df + t * t);
	const tail = 0.5 * betaI(x, df / 2, 0.5);
	return t >= 0 ? 1 - tail : tail;
};
export const tPdf = (t: number, df: number): number =>
	Math.exp(
		gammaLn((df + 1) / 2) -
			gammaLn(df / 2) -
			0.5 * Math.log(df * Math.PI) -
			((df + 1) / 2) * Math.log(1 + (t * t) / df),
	);
export const chiCdf = (x: number, df: number): number => gammaP(df / 2, x / 2);

export const positive = (...values: (number | undefined)[]): void => {
	for (const v of values) if (v === undefined || v <= 0) fail(ERR.NUM);
};
export const probability = (p: number): void => {
	if (p <= 0 || p >= 1) fail(ERR.NUM);
};

export function binomPmf(k: number, n: number, p: number): number {
	return combin(n, k) * Math.pow(p, k) * Math.pow(1 - p, n - k);
}

export function binomDist(k: number, n: number, p: number, cumulative: number): number {
	const kk = Math.trunc(k);
	const nn = Math.trunc(n);
	if (kk < 0 || kk > nn || p < 0 || p > 1) fail(ERR.NUM);
	if (!cumulative) return binomPmf(kk, nn, p);
	let total = 0;
	for (let i = 0; i <= kk; i++) total += binomPmf(i, nn, p);
	return Math.min(1, total);
}

export function poisson(k: number, mean: number, cumulative: number): number {
	const kk = Math.trunc(k);
	if (kk < 0 || mean < 0) fail(ERR.NUM);
	if (cumulative) return 1 - gammaP(kk + 1, mean);
	return Math.exp(kk * Math.log(mean) - mean - gammaLn(kk + 1));
}

export const normDist = (x: number, m: number, s: number, cumulative: number): number => {
	positive(s);
	const z = (x - m) / s;
	return cumulative ? normCdf(z) : normPdf(z) / s;
};
