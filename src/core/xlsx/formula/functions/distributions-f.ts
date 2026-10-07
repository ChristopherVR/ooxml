import { ERR, fail } from '../values';
import { betaI, binomPmf, gammaP, invert, normCdf, normInv, positive } from './dist-core';
import { numeric } from './helpers';
import { combin } from './math';
import { gammaLn } from './stats-core';
import type { FunctionSpec } from './types';

const C = 'Statistical';

const degrees = (...values: (number | undefined)[]): void => {
	for (const v of values) if (v === undefined || v < 1) fail(ERR.NUM);
};
const prob = (p: number): void => {
	if (p < 0 || p > 1) fail(ERR.NUM);
};

/** The F cumulative distribution for `x >= 0`. */
export const fCdf = (x: number, d1: number, d2: number): number =>
	x <= 0 ? 0 : betaI((d1 * x) / (d1 * x + d2), d1 / 2, d2 / 2);

function fPdf(x: number, d1: number, d2: number): number {
	if (x === 0) return d1 === 2 ? 1 : d1 > 2 ? 0 : fail(ERR.NUM);
	const lnB = gammaLn(d1 / 2) + gammaLn(d2 / 2) - gammaLn((d1 + d2) / 2);
	return Math.exp(
		(d1 / 2) * Math.log(d1) +
			(d2 / 2) * Math.log(d2) +
			(d1 / 2 - 1) * Math.log(x) -
			((d1 + d2) / 2) * Math.log(d2 + d1 * x) -
			lnB,
	);
}

const fInverse = (p: number, d1: number, d2: number): number =>
	p === 0 ? 0 : invert((x) => fCdf(x, d1, d2), p, 0, 10);

/** The beta cumulative distribution on `[lo, hi]`. */
const betaScaled = (x: number, a: number, b: number, lo: number, hi: number): number => {
	if (x < lo || x > hi || lo === hi) fail(ERR.NUM);
	return betaI((x - lo) / (hi - lo), a, b);
};

const betaInverse = (p: number, a: number, b: number, lo: number, hi: number): number => {
	positive(a, b);
	if (p <= 0 || p > 1 || lo >= hi) fail(ERR.NUM);
	return lo + (hi - lo) * invert((t) => betaI(t, a, b), p, 0, 1);
};

const gammaInverse = (p: number, a: number, b: number): number => {
	prob(p);
	positive(a, b);
	return p === 0 ? 0 : invert((x) => gammaP(a, x / b), p, 0, a * b + 1);
};

/** The smallest `k` whose cumulative binomial probability reaches `alpha`. */
function binomInverse(trials: number, p: number, alpha: number): number {
	const n = Math.trunc(trials);
	if (n < 0) fail(ERR.NUM);
	prob(p);
	prob(alpha);
	let total = 0;
	for (let k = 0; k <= n; k++) {
		total += binomPmf(k, n, p);
		if (total >= alpha) return k;
	}
	return n;
}

const hypgeomPmf = (k: number, n: number, K: number, N: number): number => {
	const kk = Math.trunc(k);
	const nn = Math.trunc(n);
	const KK = Math.trunc(K);
	const NN = Math.trunc(N);
	if (kk < 0 || nn > NN || KK > NN || kk > Math.min(nn, KK) || kk < nn - (NN - KK)) fail(ERR.NUM);
	return (combin(KK, kk) * combin(NN - KK, nn - kk)) / combin(NN, nn);
};

const negbinomPmf = (f: number, s: number, p: number): number => {
	const ff = Math.trunc(f);
	const ss = Math.trunc(s);
	if (ff < 0 || ss < 1 || p < 0 || p > 1) fail(ERR.NUM);
	return combin(ff + ss - 1, ss - 1) * Math.pow(p, ss) * Math.pow(1 - p, ff);
};

/** F, beta and gamma inverses, plus the pre-2010 names of the distribution functions. */
export const DISTRIBUTION_F: FunctionSpec[] = [
	numeric(
		'F.DIST',
		C,
		'F.DIST(x, deg_freedom1, deg_freedom2, cumulative)',
		'The left-tailed F probability distribution.',
		4,
		4,
		(x, d1, d2, c) => {
			degrees(d1, d2);
			if (x < 0) fail(ERR.NUM);
			const a = Math.trunc(d1 ?? 1);
			const b = Math.trunc(d2 ?? 1);
			return c ? fCdf(x, a, b) : fPdf(x, a, b);
		},
	),
	...(['F.DIST.RT', 'FDIST'] as const).map((name) =>
		numeric(
			name,
			C,
			`${name}(x, deg_freedom1, deg_freedom2)`,
			'The right-tailed F probability distribution.',
			3,
			3,
			(x, d1, d2) => {
				degrees(d1, d2);
				if (x < 0) fail(ERR.NUM);
				return 1 - fCdf(x, Math.trunc(d1 ?? 1), Math.trunc(d2 ?? 1));
			},
		),
	),
	numeric(
		'F.INV',
		C,
		'F.INV(probability, deg_freedom1, deg_freedom2)',
		'The inverse of the left-tailed F distribution.',
		3,
		3,
		(p, d1, d2) => {
			degrees(d1, d2);
			if (p < 0 || p >= 1) fail(ERR.NUM);
			return fInverse(p, Math.trunc(d1 ?? 1), Math.trunc(d2 ?? 1));
		},
	),
	...(['F.INV.RT', 'FINV'] as const).map((name) =>
		numeric(
			name,
			C,
			`${name}(probability, deg_freedom1, deg_freedom2)`,
			'The inverse of the right-tailed F distribution.',
			3,
			3,
			(p, d1, d2) => {
				degrees(d1, d2);
				if (p <= 0 || p > 1) fail(ERR.NUM);
				return fInverse(1 - p, Math.trunc(d1 ?? 1), Math.trunc(d2 ?? 1));
			},
		),
	),
	...(['BETA.INV', 'BETAINV'] as const).map((name) =>
		numeric(
			name,
			C,
			`${name}(probability, alpha, beta, [A], [B])`,
			'The inverse of the cumulative beta distribution.',
			3,
			5,
			(p, a, b, lo, hi) => betaInverse(p, a ?? 1, b ?? 1, lo ?? 0, hi ?? 1),
			[0, 0, 0, 0, 1],
		),
	),
	numeric(
		'BETADIST',
		C,
		'BETADIST(x, alpha, beta, [A], [B])',
		'The cumulative beta distribution.',
		3,
		5,
		(x, a, b, lo, hi) => {
			positive(a, b);
			return betaScaled(x, a ?? 1, b ?? 1, lo ?? 0, hi ?? 1);
		},
		[0, 0, 0, 0, 1],
	),
	...(['GAMMA.INV', 'GAMMAINV'] as const).map((name) =>
		numeric(
			name,
			C,
			`${name}(probability, alpha, beta)`,
			'The inverse of the gamma cumulative distribution.',
			3,
			3,
			(p, a, b) => gammaInverse(p, a ?? 1, b ?? 1),
		),
	),
	numeric(
		'GAMMADIST',
		C,
		'GAMMADIST(x, alpha, beta, cumulative)',
		'The gamma distribution.',
		4,
		4,
		(x, a, b, c) => {
			if (x < 0) fail(ERR.NUM);
			positive(a, b);
			const al = a ?? 1;
			const be = b ?? 1;
			return c
				? gammaP(al, x / be)
				: Math.exp((al - 1) * Math.log(x) - x / be - gammaLn(al) - al * Math.log(be));
		},
	),
	numeric(
		'LOGINV',
		C,
		'LOGINV(probability, mean, standard_dev)',
		'The inverse of the lognormal cumulative distribution.',
		3,
		3,
		(p, m, s) => {
			positive(s);
			if (p <= 0 || p >= 1) fail(ERR.NUM);
			return Math.exp((m ?? 0) + (s ?? 1) * normInv(p));
		},
	),
	numeric(
		'LOGNORMDIST',
		C,
		'LOGNORMDIST(x, mean, standard_dev)',
		'The cumulative lognormal distribution.',
		3,
		3,
		(x, m, s) => {
			positive(x, s);
			return normCdf((Math.log(x) - (m ?? 0)) / (s ?? 1));
		},
	),
	...(['BINOM.INV', 'CRITBINOM'] as const).map((name) =>
		numeric(
			name,
			C,
			`${name}(trials, probability_s, alpha)`,
			'The smallest value whose cumulative binomial distribution reaches the criterion.',
			3,
			3,
			(n, p, alpha) => binomInverse(n, p ?? 0, alpha ?? 0),
		),
	),
	numeric(
		'BINOM.DIST.RANGE',
		C,
		'BINOM.DIST.RANGE(trials, probability_s, number_s, [number_s2])',
		'The probability of a trial result within a range, from the binomial distribution.',
		3,
		4,
		(trials, p, s, s2) => {
			const n = Math.trunc(trials);
			const lo = Math.trunc(s ?? 0);
			const hi = Math.trunc(s2 ?? s ?? 0);
			prob(p ?? 0);
			if (n < 0 || lo < 0 || lo > n || hi < lo || hi > n) fail(ERR.NUM);
			let total = 0;
			for (let k = lo; k <= hi; k++) total += binomPmf(k, n, p ?? 0);
			return Math.min(1, total);
		},
	),
	numeric(
		'HYPGEOMDIST',
		C,
		'HYPGEOMDIST(sample_s, number_sample, population_s, number_pop)',
		'The hypergeometric distribution.',
		4,
		4,
		hypgeomPmf,
	),
	numeric(
		'NEGBINOMDIST',
		C,
		'NEGBINOMDIST(number_f, number_s, probability_s)',
		'The negative binomial distribution.',
		3,
		3,
		negbinomPmf,
	),
];
