import { ERR, fail, type Scalar } from '../values.js';
import { criteriaPairs, liftCriteria, matchingValues } from './criteria.js';
import { all, collectNumbers, num, spec } from './helpers.js';
import { STATISTICAL_MORE } from './statistical-more.js';
import * as S from './stats-core.js';
import type { FunctionSpec } from './types.js';
import { STATISTICAL_SPREAD } from './statistical-spread.js';
import { COUNT_FUNCTIONS } from './statistical-count.js';

const C = 'Statistical';

const numbersOf = (values: Scalar[]): number[] =>
	values.filter((v): v is number => typeof v === 'number');

const extreme = S.extreme;

export const STATISTICAL_FUNCTIONS: FunctionSpec[] = [
	...STATISTICAL_SPREAD,
	...STATISTICAL_MORE,
	...COUNT_FUNCTIONS,
	spec(
		'AVERAGEA',
		C,
		'AVERAGEA(value1, [value2], ...)',
		'The mean counting text as 0 and logicals.',
		1,
		255,
		(args, ctx) => S.mean(all(ctx, args)),
		['any'],
	),
	spec(
		'AVERAGEIF',
		C,
		'AVERAGEIF(range, criteria, [average_range])',
		'The mean of the cells that meet a criterion.',
		2,
		3,
		(args, ctx) =>
			S.mean(
				numbersOf(
					matchingValues(
						ctx,
						[{ range: args[0] ?? null, criteria: args[1] ?? null }],
						args[2] ?? undefined,
						false,
					),
				),
			),
		['any', 'value', 'any'],
	),
	spec(
		'AVERAGEIFS',
		C,
		'AVERAGEIFS(average_range, criteria_range1, criteria1, ...)',
		'The mean of the cells that meet several criteria.',
		3,
		255,
		(args, ctx) =>
			liftCriteria(ctx, criteriaPairs(args, 1), (pairs) =>
				S.mean(numbersOf(matchingValues(ctx, pairs, args[0] ?? null, true))),
			),
		['any'],
	),
	spec(
		'MAX',
		C,
		'MAX(number1, [number2], ...)',
		'The largest value.',
		1,
		255,
		(args, ctx) => extreme(collectNumbers(ctx, args), true),
		['any'],
	),
	spec(
		'MIN',
		C,
		'MIN(number1, [number2], ...)',
		'The smallest value.',
		1,
		255,
		(args, ctx) => extreme(collectNumbers(ctx, args), false),
		['any'],
	),
	spec(
		'MAXA',
		C,
		'MAXA(value1, [value2], ...)',
		'The largest value counting logicals and text.',
		1,
		255,
		(args, ctx) => extreme(all(ctx, args), true),
		['any'],
	),
	spec(
		'MINA',
		C,
		'MINA(value1, [value2], ...)',
		'The smallest value counting logicals and text.',
		1,
		255,
		(args, ctx) => extreme(all(ctx, args), false),
		['any'],
	),
	spec(
		'MAXIFS',
		C,
		'MAXIFS(max_range, criteria_range1, criteria1, ...)',
		'The largest value meeting criteria.',
		3,
		255,
		(args, ctx) =>
			liftCriteria(ctx, criteriaPairs(args, 1), (pairs) =>
				extreme(numbersOf(matchingValues(ctx, pairs, args[0] ?? null, true)), true),
			),
		['any'],
	),
	spec(
		'MINIFS',
		C,
		'MINIFS(min_range, criteria_range1, criteria1, ...)',
		'The smallest value meeting criteria.',
		3,
		255,
		(args, ctx) =>
			liftCriteria(ctx, criteriaPairs(args, 1), (pairs) =>
				extreme(numbersOf(matchingValues(ctx, pairs, args[0] ?? null, true)), false),
			),
		['any'],
	),
	spec(
		'MEDIAN',
		C,
		'MEDIAN(number1, [number2], ...)',
		'The median.',
		1,
		255,
		(args, ctx) => S.median(collectNumbers(ctx, args)),
		['any'],
	),
	spec(
		'MODE',
		C,
		'MODE(number1, [number2], ...)',
		'The most frequent value.',
		1,
		255,
		(args, ctx) => S.modes(collectNumbers(ctx, args))[0] ?? fail(ERR.NA),
		['any'],
	),
	spec(
		'MODE.SNGL',
		C,
		'MODE.SNGL(number1, [number2], ...)',
		'The most frequent value.',
		1,
		255,
		(args, ctx) => S.modes(collectNumbers(ctx, args))[0] ?? fail(ERR.NA),
		['any'],
	),
	spec(
		'LARGE',
		C,
		'LARGE(array, k)',
		'The k-th largest value.',
		2,
		2,
		(args, ctx) => S.kth(collectNumbers(ctx, [args[0] ?? null]), num(args[1]), true),
		['any', 'value'],
	),
	spec(
		'SMALL',
		C,
		'SMALL(array, k)',
		'The k-th smallest value.',
		2,
		2,
		(args, ctx) => S.kth(collectNumbers(ctx, [args[0] ?? null]), num(args[1]), false),
		['any', 'value'],
	),
	...(['RANK', 'RANK.EQ', 'RANK.AVG'] as const).map((name) =>
		spec(
			name,
			C,
			`${name}(number, ref, [order])`,
			'The rank of a number in a list.',
			2,
			3,
			(args, ctx) => {
				const x = num(args[0]);
				const values = collectNumbers(ctx, [args[1] ?? null]);
				const ascending = args.length > 2 && num(args[2]) !== 0;
				return S.rank(values, x, ascending, name === 'RANK.AVG');
			},
			['value', 'any', 'value'],
		),
	),
];
