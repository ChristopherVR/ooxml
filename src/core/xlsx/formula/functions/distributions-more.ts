import { ERR, fail } from '../values.js';
import { betaI, chiCdf, invert, positive, probability, tCdf, tPdf } from './dist-core.js';
import { numeric } from './helpers.js';
import { combin } from './math.js';
import { gammaLn } from './stats-core.js';
import type { FunctionSpec } from './types.js';

const C = 'Statistical';

/** Student's t, chi-squared, beta and discrete distributions. */
export const DISTRIBUTION_MORE: FunctionSpec[] = [
	numeric(
		'T.DIST',
		C,
		'T.DIST(x, deg_freedom, cumulative)',
		"Student's left-tailed t-distribution.",
		3,
		3,
		(x, df, c) => {
			if ((df ?? 0) < 1) fail(ERR.NUM);
			return c ? tCdf(x, Math.trunc(df ?? 1)) : tPdf(x, Math.trunc(df ?? 1));
		},
	),
	numeric(
		'T.DIST.RT',
		C,
		'T.DIST.RT(x, deg_freedom)',
		"Student's right-tailed t-distribution.",
		2,
		2,
		(x, df) => {
			if ((df ?? 0) < 1) fail(ERR.NUM);
			return 1 - tCdf(x, Math.trunc(df ?? 1));
		},
	),
	numeric(
		'T.DIST.2T',
		C,
		'T.DIST.2T(x, deg_freedom)',
		"Student's two-tailed t-distribution.",
		2,
		2,
		(x, df) => {
			if (x < 0 || (df ?? 0) < 1) fail(ERR.NUM);
			return 2 * (1 - tCdf(x, Math.trunc(df ?? 1)));
		},
	),
	numeric(
		'TDIST',
		C,
		'TDIST(x, deg_freedom, tails)',
		"Student's t-distribution.",
		3,
		3,
		(x, df, tails) => {
			if (x < 0 || (df ?? 0) < 1 || (tails !== 1 && tails !== 2)) fail(ERR.NUM);
			return (tails ?? 1) * (1 - tCdf(x, Math.trunc(df ?? 1)));
		},
	),
	numeric(
		'T.INV',
		C,
		'T.INV(probability, deg_freedom)',
		"The left-tailed inverse of Student's t-distribution.",
		2,
		2,
		(p, df) => {
			probability(p);
			if ((df ?? 0) < 1) fail(ERR.NUM);
			const d = Math.trunc(df ?? 1);
			return p < 0.5
				? -invert((t) => tCdf(t, d), 1 - p, 0, 10)
				: invert((t) => tCdf(t, d), p, 0, 10);
		},
	),
	...(['T.INV.2T', 'TINV'] as const).map((name) =>
		numeric(
			name,
			C,
			`${name}(probability, deg_freedom)`,
			"The two-tailed inverse of Student's t-distribution.",
			2,
			2,
			(p, df) => {
				if (p <= 0 || p > 1 || (df ?? 0) < 1) fail(ERR.NUM);
				return invert((t) => tCdf(t, Math.trunc(df ?? 1)), 1 - p / 2, 0, 10);
			},
		),
	),
	numeric(
		'CHISQ.DIST',
		C,
		'CHISQ.DIST(x, deg_freedom, cumulative)',
		'The chi-squared distribution.',
		3,
		3,
		(x, df, c) => {
			if (x < 0 || (df ?? 0) < 1) fail(ERR.NUM);
			const k = Math.trunc(df ?? 1);
			return c
				? chiCdf(x, k)
				: Math.exp((k / 2 - 1) * Math.log(x) - x / 2 - (k / 2) * Math.LN2 - gammaLn(k / 2));
		},
	),
	...(['CHISQ.DIST.RT', 'CHIDIST'] as const).map((name) =>
		numeric(
			name,
			C,
			`${name}(x, deg_freedom)`,
			'The right-tailed chi-squared probability.',
			2,
			2,
			(x, df) => {
				if (x < 0 || (df ?? 0) < 1) fail(ERR.NUM);
				return 1 - chiCdf(x, Math.trunc(df ?? 1));
			},
		),
	),
	numeric(
		'CHISQ.INV',
		C,
		'CHISQ.INV(probability, deg_freedom)',
		'The inverse of the left-tailed chi-squared distribution.',
		2,
		2,
		(p, df) => {
			if (p < 0 || p >= 1 || (df ?? 0) < 1) fail(ERR.NUM);
			return invert((x) => chiCdf(x, Math.trunc(df ?? 1)), p, 0, 100);
		},
	),
	...(['CHISQ.INV.RT', 'CHIINV'] as const).map((name) =>
		numeric(
			name,
			C,
			`${name}(probability, deg_freedom)`,
			'The inverse of the right-tailed chi-squared distribution.',
			2,
			2,
			(p, df) => {
				if (p <= 0 || p > 1 || (df ?? 0) < 1) fail(ERR.NUM);
				return invert((x) => chiCdf(x, Math.trunc(df ?? 1)), 1 - p, 0, 100);
			},
		),
	),
	numeric(
		'BETA.DIST',
		C,
		'BETA.DIST(x, alpha, beta, cumulative, [A], [B])',
		'The beta distribution.',
		4,
		6,
		(x, a, b, c, lo, hi) => {
			const A = lo ?? 0;
			const B = hi ?? 1;
			positive(a, b);
			if (x < A || x > B || A === B) fail(ERR.NUM);
			const t = (x - A) / (B - A);
			const al = a ?? 1;
			const be = b ?? 1;
			if (c) return betaI(t, al, be);
			const lnB = gammaLn(al) + gammaLn(be) - gammaLn(al + be);
			return Math.exp((al - 1) * Math.log(t) + (be - 1) * Math.log(1 - t) - lnB) / (B - A);
		},
	),
	numeric(
		'NEGBINOM.DIST',
		C,
		'NEGBINOM.DIST(number_f, number_s, probability_s, cumulative)',
		'The negative binomial distribution.',
		4,
		4,
		(f, s, p, c) => {
			const ff = Math.trunc(f);
			const ss = Math.trunc(s ?? 0);
			const pp = p ?? 0;
			if (ff < 0 || ss < 1 || pp < 0 || pp > 1) fail(ERR.NUM);
			const pmf = (k: number): number =>
				combin(k + ss - 1, ss - 1) * Math.pow(pp, ss) * Math.pow(1 - pp, k);
			if (!c) return pmf(ff);
			let total = 0;
			for (let k = 0; k <= ff; k++) total += pmf(k);
			return total;
		},
	),
	numeric(
		'HYPGEOM.DIST',
		C,
		'HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, cumulative)',
		'The hypergeometric distribution.',
		5,
		5,
		(k, n, K, N, c) => {
			const kk = Math.trunc(k);
			const nn = Math.trunc(n ?? 0);
			const KK = Math.trunc(K ?? 0);
			const NN = Math.trunc(N ?? 0);
			if (kk < 0 || nn > NN || KK > NN || kk > Math.min(nn, KK)) fail(ERR.NUM);
			const pmf = (i: number): number => (combin(KK, i) * combin(NN - KK, nn - i)) / combin(NN, nn);
			if (!c) return pmf(kk);
			let total = 0;
			for (let i = Math.max(0, nn - (NN - KK)); i <= kk; i++) total += pmf(i);
			return total;
		},
	),
];
