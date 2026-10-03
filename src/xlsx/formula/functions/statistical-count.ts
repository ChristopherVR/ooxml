// COUNT, COUNTA, COUNTBLANK, COUNTIF and COUNTIFS.
import { parseNumberText } from '../text-number.js';
import { RefValue, type Value } from '../values.js';
import type { CallContext } from '../context.js';
import { criteriaPairs, liftCriteria, matchingCells } from './criteria.js';
import { collectNumbers, makeCriteria, scalar, spec } from './helpers.js';
import * as S from './stats-core.js';
import type { FunctionSpec } from './types.js';

const C = 'Statistical';

function count(ctx: CallContext, args: Value[]): number {
	let n = 0;
	for (const arg of args) {
		ctx.forEach(arg, (value, kind) => {
			if (typeof value === 'number') n++;
			else if (kind === 'direct') {
				// Omitted arguments count as zeros: COUNT(,) is 2.
				if (typeof value === 'boolean' || value === null) n++;
				else if (typeof value === 'string' && parseNumberText(value) !== undefined) n++;
			}
		});
	}
	return n;
}

function countA(ctx: CallContext, args: Value[]): number {
	let n = 0;
	for (const arg of args) {
		ctx.forEach(arg, (value, kind) => {
			if (value !== null || kind === 'direct') n++;
		});
	}
	return n;
}

function cellsIn(value: Value): number {
	if (value instanceof RefValue) {
		return value.areas.reduce(
			(acc, a) =>
				acc + (a.range.end.row - a.range.start.row + 1) * (a.range.end.col - a.range.start.col + 1),
			0,
		);
	}
	return 1;
}

function countBlank(ctx: CallContext, value: Value): number {
	if (!(value instanceof RefValue)) {
		return ctx
			.toMatrix(value)
			.flat()
			.filter((v) => v === null || v === '').length;
	}
	let nonBlank = 0;
	ctx.forEach(value, (v) => {
		if (v !== null && v !== '') nonBlank++;
	});
	return cellsIn(value) - nonBlank;
}

function countIf(ctx: CallContext, range: Value, criteria: Value): number {
	if (range instanceof RefValue && range.areas.length === 1) {
		// One area: read through the shared (cached) criteria path, blanks past the used part counted.
		const found = matchingCells(ctx, [{ range, criteria }], undefined, false);
		return found.values.length + found.blankTail;
	}
	const test = makeCriteria(scalar(criteria));
	if (!(range instanceof RefValue)) {
		let n = 0;
		ctx.toMatrix(range).forEachValue((v) => {
			if (test(v)) n++;
		});
		return n;
	}
	let matched = 0;
	let stored = 0;
	ctx.forEach(range, (v) => {
		stored++;
		if (test(v)) matched++;
	});
	return test(null) ? matched + cellsIn(range) - stored : matched;
}

export const COUNT_FUNCTIONS: FunctionSpec[] = [
	spec(
		'COUNT',
		C,
		'COUNT(value1, [value2], ...)',
		'Counts the numbers in the arguments.',
		0,
		255,
		(args, ctx) => count(ctx, args),
		['any'],
	),
	spec(
		'COUNTA',
		C,
		'COUNTA(value1, [value2], ...)',
		'Counts the non-empty values.',
		1,
		255,
		(args, ctx) => countA(ctx, args),
		['any'],
	),
	spec(
		'COUNTBLANK',
		C,
		'COUNTBLANK(range)',
		'Counts the empty cells in a range.',
		1,
		1,
		(args, ctx) => countBlank(ctx, args[0] ?? null),
		['any'],
	),
	spec(
		'COUNTIF',
		C,
		'COUNTIF(range, criteria)',
		'Counts the cells that meet a criterion.',
		2,
		2,
		(args, ctx) => countIf(ctx, args[0] ?? null, args[1] ?? null),
		['any', 'value'],
	),
	spec(
		'COUNTIFS',
		C,
		'COUNTIFS(criteria_range1, criteria1, ...)',
		'Counts the cells that meet several criteria.',
		2,
		254,
		(args, ctx) =>
			liftCriteria(ctx, criteriaPairs(args, 0), (pairs) => {
				const found = matchingCells(ctx, pairs, undefined, true);
				return found.values.length + found.blankTail;
			}),
		['any'],
	),
	spec(
		'AVERAGE',
		C,
		'AVERAGE(number1, [number2], ...)',
		'The arithmetic mean.',
		1,
		255,
		(args, ctx) => S.mean(collectNumbers(ctx, args)),
		['any'],
	),
];
