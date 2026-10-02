import { ERR, fail } from '../values.js';
import {
	binomDist,
	erf,
	gammaP,
	invert,
	normCdf,
	normDist,
	normInv,
	normPdf,
	poisson,
	positive,
	probability,
	tCdf,
} from './dist-core.js';
import { numeric } from './helpers.js';
import { gammaLn } from './stats-core.js';
import type { FunctionSpec } from './types.js';
import { DISTRIBUTION_MORE } from './distributions-more.js';

const C = 'Statistical';

export const DISTRIBUTION_FUNCTIONS: FunctionSpec[] = [
	...DISTRIBUTION_MORE,
	numeric(
		'NORM.DIST',
		C,
		'NORM.DIST(x, mean, standard_dev, cumulative)',
		'The normal distribution.',
		4,
		4,
		(x, m, s, c) => normDist(x, m ?? 0, s ?? 1, c ?? 0),
	),
	numeric(
		'NORMDIST',
		C,
		'NORMDIST(x, mean, standard_dev, cumulative)',
		'The normal distribution.',
		4,
		4,
		(x, m, s, c) => normDist(x, m ?? 0, s ?? 1, c ?? 0),
	),
	numeric(
		'NORM.S.DIST',
		C,
		'NORM.S.DIST(z, cumulative)',
		'The standard normal distribution.',
		2,
		2,
		(z, c) => (c ? normCdf(z) : normPdf(z)),
	),
	numeric(
		'NORMSDIST',
		C,
		'NORMSDIST(z)',
		'The standard normal cumulative distribution.',
		1,
		1,
		normCdf,
	),
	numeric(
		'NORM.INV',
		C,
		'NORM.INV(probability, mean, standard_dev)',
		'The inverse of the normal cumulative distribution.',
		3,
		3,
		(p, m, s) => {
			positive(s);
			return (m ?? 0) + (s ?? 1) * normInv(p);
		},
	),
	numeric(
		'NORMINV',
		C,
		'NORMINV(probability, mean, standard_dev)',
		'The inverse of the normal cumulative distribution.',
		3,
		3,
		(p, m, s) => {
			positive(s);
			return (m ?? 0) + (s ?? 1) * normInv(p);
		},
	),
	numeric(
		'NORM.S.INV',
		C,
		'NORM.S.INV(probability)',
		'The inverse of the standard normal distribution.',
		1,
		1,
		normInv,
	),
	numeric(
		'NORMSINV',
		C,
		'NORMSINV(probability)',
		'The inverse of the standard normal distribution.',
		1,
		1,
		normInv,
	),
	numeric('PHI', C, 'PHI(x)', 'The standard normal density.', 1, 1, normPdf),
	numeric(
		'GAUSS',
		C,
		'GAUSS(z)',
		'The probability between the mean and z standard deviations.',
		1,
		1,
		(z) => normCdf(z) - 0.5,
	),
	numeric(
		'ERF',
		'Engineering',
		'ERF(lower_limit, [upper_limit])',
		'The error function.',
		1,
		2,
		(a, b) => (b === undefined ? erf(a) : erf(b) - erf(a)),
	),
	numeric('ERF.PRECISE', 'Engineering', 'ERF.PRECISE(x)', 'The error function.', 1, 1, erf),
	numeric(
		'ERFC',
		'Engineering',
		'ERFC(x)',
		'The complementary error function.',
		1,
		1,
		(x) => 1 - erf(x),
	),
	numeric(
		'ERFC.PRECISE',
		'Engineering',
		'ERFC.PRECISE(x)',
		'The complementary error function.',
		1,
		1,
		(x) => 1 - erf(x),
	),
	...(['CONFIDENCE', 'CONFIDENCE.NORM'] as const).map((name) =>
		numeric(
			name,
			C,
			`${name}(alpha, standard_dev, size)`,
			'The confidence interval for a population mean (normal).',
			3,
			3,
			(a, s, n) => {
				probability(a);
				positive(s);
				if ((n ?? 0) < 1) fail(ERR.NUM);
				return (normInv(1 - a / 2) * (s ?? 1)) / Math.sqrt(Math.trunc(n ?? 1));
			},
		),
	),
	numeric(
		'CONFIDENCE.T',
		C,
		'CONFIDENCE.T(alpha, standard_dev, size)',
		"The confidence interval for a population mean (Student's t).",
		3,
		3,
		(a, s, n) => {
			probability(a);
			positive(s);
			const size = Math.trunc(n ?? 1);
			if (size < 2) fail(ERR.DIV0);
			return (invert((t) => tCdf(t, size - 1), 1 - a / 2, 0, 10) * (s ?? 1)) / Math.sqrt(size);
		},
	),
	...(['BINOM.DIST', 'BINOMDIST'] as const).map((name) =>
		numeric(
			name,
			C,
			`${name}(number_s, trials, probability_s, cumulative)`,
			'The binomial distribution probability.',
			4,
			4,
			(k, n, p, c) => binomDist(k, n ?? 0, p ?? 0, c ?? 0),
		),
	),
	...(['POISSON.DIST', 'POISSON'] as const).map((name) =>
		numeric(name, C, `${name}(x, mean, cumulative)`, 'The Poisson distribution.', 3, 3, (k, m, c) =>
			poisson(k, m ?? 0, c ?? 0),
		),
	),
	...(['EXPON.DIST', 'EXPONDIST'] as const).map((name) =>
		numeric(
			name,
			C,
			`${name}(x, lambda, cumulative)`,
			'The exponential distribution.',
			3,
			3,
			(x, l, c) => {
				if (x < 0) fail(ERR.NUM);
				positive(l);
				return c ? 1 - Math.exp(-(l ?? 1) * x) : (l ?? 1) * Math.exp(-(l ?? 1) * x);
			},
		),
	),
	numeric(
		'LOGNORM.DIST',
		C,
		'LOGNORM.DIST(x, mean, standard_dev, cumulative)',
		'The lognormal distribution.',
		4,
		4,
		(x, m, s, c) => {
			positive(x, s);
			const z = (Math.log(x) - (m ?? 0)) / (s ?? 1);
			return c ? normCdf(z) : normPdf(z) / (x * (s ?? 1));
		},
	),
	numeric(
		'LOGNORM.INV',
		C,
		'LOGNORM.INV(probability, mean, standard_dev)',
		'The inverse of the lognormal distribution.',
		3,
		3,
		(p, m, s) => {
			positive(s);
			return Math.exp((m ?? 0) + (s ?? 1) * normInv(p));
		},
	),
	...(['WEIBULL.DIST', 'WEIBULL'] as const).map((name) =>
		numeric(
			name,
			C,
			`${name}(x, alpha, beta, cumulative)`,
			'The Weibull distribution.',
			4,
			4,
			(x, a, b, c) => {
				if (x < 0) fail(ERR.NUM);
				positive(a, b);
				const al = a ?? 1;
				const be = b ?? 1;
				return c
					? 1 - Math.exp(-Math.pow(x / be, al))
					: (al / Math.pow(be, al)) * Math.pow(x, al - 1) * Math.exp(-Math.pow(x / be, al));
			},
		),
	),
	numeric(
		'GAMMA.DIST',
		C,
		'GAMMA.DIST(x, alpha, beta, cumulative)',
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
];
