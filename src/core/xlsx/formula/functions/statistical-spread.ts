import { collectNumbers, num, spec } from './helpers.js';
import { all } from './helpers.js';
import * as S from './stats-core.js';
import type { FunctionSpec } from './types.js';

const C = 'Statistical';

/** Dispersion, percentiles and ranks. */
export const STATISTICAL_SPREAD: FunctionSpec[] = [
	...(
		[
			['STDEV', true, false],
			['STDEV.S', true, false],
			['STDEVP', false, false],
			['STDEV.P', false, false],
			['STDEVA', true, true],
			['STDEVPA', false, true],
		] as const
	).map(([name, sample, a]) =>
		spec(
			name,
			C,
			`${name}(number1, [number2], ...)`,
			`The ${sample ? 'sample' : 'population'} standard deviation${a ? ' counting logicals and text' : ''}.`,
			1,
			255,
			(args, ctx) => S.stdev(a ? all(ctx, args) : collectNumbers(ctx, args), sample),
			['any'],
		),
	),
	...(
		[
			['VAR', true, false],
			['VAR.S', true, false],
			['VARP', false, false],
			['VAR.P', false, false],
			['VARA', true, true],
			['VARPA', false, true],
		] as const
	).map(([name, sample, a]) =>
		spec(
			name,
			C,
			`${name}(number1, [number2], ...)`,
			`The ${sample ? 'sample' : 'population'} variance${a ? ' counting logicals and text' : ''}.`,
			1,
			255,
			(args, ctx) => S.variance(a ? all(ctx, args) : collectNumbers(ctx, args), sample),
			['any'],
		),
	),
	...(['PERCENTILE', 'PERCENTILE.INC', 'PERCENTILE.EXC'] as const).map((name) =>
		spec(
			name,
			C,
			`${name}(array, k)`,
			'The k-th percentile.',
			2,
			2,
			(args, ctx) => {
				const values = collectNumbers(ctx, [args[0] ?? null]);
				const k = num(args[1]);
				return name === 'PERCENTILE.EXC' ? S.percentileExc(values, k) : S.percentileInc(values, k);
			},
			['any', 'value'],
		),
	),
	...(['QUARTILE', 'QUARTILE.INC', 'QUARTILE.EXC'] as const).map((name) =>
		spec(
			name,
			C,
			`${name}(array, quart)`,
			'A quartile of a data set.',
			2,
			2,
			(args, ctx) =>
				S.quartile(collectNumbers(ctx, [args[0] ?? null]), num(args[1]), name === 'QUARTILE.EXC'),
			['any', 'value'],
		),
	),
	...(['PERCENTRANK', 'PERCENTRANK.INC', 'PERCENTRANK.EXC'] as const).map((name) =>
		spec(
			name,
			C,
			`${name}(array, x, [significance])`,
			'The rank of a value as a percentage of the data set.',
			2,
			3,
			(args, ctx) =>
				S.percentRank(
					collectNumbers(ctx, [args[0] ?? null]),
					num(args[1]),
					args.length > 2 ? num(args[2]) : 3,
					name === 'PERCENTRANK.EXC',
				),
			['any', 'value', 'value'],
		),
	),
];
