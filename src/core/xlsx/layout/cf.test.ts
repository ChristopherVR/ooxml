import { describe, expect, it } from 'vitest';
import type { CellRange } from '../address.js';
import { getValue, putCell } from '../cells.js';
import type { CellValue, ConditionalRule, DifferentialStyle, Workbook } from '../model.js';
import { createWorkbook } from '../workbook.js';
import {
	createConditionalFormatEvaluator,
	scaleColor,
	type FormulaEvaluator,
} from './cf-evaluator.js';
import { compareValues, percentile, testOperator } from './cf-values.js';

const range = (r1: number, c1: number, r2: number, c2: number): CellRange => ({
	start: { row: r1, col: c1 },
	end: { row: r2, col: c2 },
});

const RED: DifferentialStyle = { font: { color: { rgb: '9C0006' } } };
const GREEN: DifferentialStyle = {
	fill: { type: 'pattern', pattern: 'solid', bgColor: { rgb: 'C6EFCE' } },
};
const BOLD: DifferentialStyle = { font: { bold: true } };

/** A workbook whose column A holds `values` from row 0. */
function column(values: CellValue[], rules: ConditionalRule[], to = values.length - 1): Workbook {
	const wb = createWorkbook();
	const sheet = wb.sheets[0]!;
	values.forEach((v, i) => {
		if (v !== null) putCell(sheet, i, 0, { value: v });
	});
	sheet.conditionalFormats.push({ ranges: [range(0, 0, to, 0)], rules });
	return wb;
}

/** Evaluates `A<n> op k` style formulas the tests use, reading the live sheet. */
const tinyEvaluate =
	(wb: Workbook): FormulaEvaluator =>
	(formula) => {
		const m = /^\$?A\$?(\d+)\s*(>|<|=)\s*(-?\d+)$/.exec(formula);
		if (!m) return { error: '#NAME?' };
		const v = getValue(wb.sheets[0]!, Number(m[1]) - 1, 0);
		const k = Number(m[3]);
		if (typeof v !== 'number') return false;
		return m[2] === '>' ? v > k : m[2] === '<' ? v < k : v === k;
	};

/** Moves `A<n>` row references (enough for these tests; the real one comes from the formula module). */
const translate = (formula: string, dRow: number): string =>
	formula.replace(/(\$?)A(\$?)(\d+)/g, (_, c, r, n) => `${c}A${r}${r ? n : Number(n) + dRow}`);

const evaluator = (wb: Workbook) =>
	createConditionalFormatEvaluator(wb, 0, tinyEvaluate(wb), { translate });

describe('Excel value comparison', () => {
	it('orders numbers < text < booleans and compares text case-insensitively', () => {
		expect(compareValues(1, 'a')).toBeLessThan(0);
		expect(compareValues('a', true)).toBeLessThan(0);
		expect(compareValues('ABC', 'abc')).toBe(0);
		expect(compareValues(null, 0)).toBe(0);
		expect(compareValues(null, '')).toBe(0);
		expect(compareValues(2, 10)).toBeLessThan(0);
	});

	it('tests cellIs operators', () => {
		expect(testOperator('between', 5, 1, 10)).toBe(true);
		expect(testOperator('between', 5, 10, 1)).toBe(true);
		expect(testOperator('notBetween', 11, 1, 10)).toBe(true);
		expect(testOperator('greaterThanOrEqual', 5, 5, null)).toBe(true);
		expect(testOperator('notEqual', 'x', 'X', null)).toBe(false);
		expect(testOperator('lessThan', null, 5, null)).toBe(true);
		expect(testOperator('equal', { error: '#N/A' }, 1, null)).toBe(false);
	});

	it('computes PERCENTILE.INC', () => {
		expect(percentile([1, 2, 3, 4, 5], 0.5)).toBe(3);
		expect(percentile([1, 2, 3, 4], 0.25)).toBe(1.75);
		expect(percentile([], 0.5)).toBe(0);
	});
});

describe('conditional format rules', () => {
	it('evaluates cellIs with literal operands', () => {
		const wb = column(
			[1, 5, 10, 'x'],
			[{ type: 'cellIs', operator: 'greaterThan', formulas: ['4'], style: RED, priority: 1 }],
		);
		const cf = evaluator(wb);
		expect(cf.at(0, 0)).toBeUndefined();
		expect(cf.at(1, 0)?.style).toEqual(RED);
		expect(cf.at(2, 0)?.style).toEqual(RED);
		// Text is greater than any number in Excel's ordering.
		expect(cf.at(3, 0)?.style).toEqual(RED);
		expect(cf.at(0, 1)).toBeUndefined();
	});

	it('evaluates cellIs between and string equality', () => {
		const wb = column(
			[1, 5, 'Yes'],
			[
				{ type: 'cellIs', operator: 'between', formulas: ['2', '6'], style: RED, priority: 1 },
				{ type: 'cellIs', operator: 'equal', formulas: ['"yes"'], style: BOLD, priority: 2 },
			],
		);
		const cf = evaluator(wb);
		expect(cf.at(0, 0)).toBeUndefined();
		expect(cf.at(1, 0)?.style).toEqual(RED);
		expect(cf.at(2, 0)?.style).toEqual(BOLD);
	});

	it('evaluates expressions relative to the range top-left', () => {
		const wb = column(
			[1, 5, 10],
			[{ type: 'expression', formula: 'A1>4', style: GREEN, priority: 1 }],
		);
		const cf = evaluator(wb);
		expect(cf.at(0, 0)).toBeUndefined();
		expect(cf.at(1, 0)?.style).toEqual(GREEN);
		expect(cf.at(2, 0)?.style).toEqual(GREEN);
	});

	it('keeps absolute references fixed', () => {
		const wb = column(
			[5, 1, 1],
			[{ type: 'expression', formula: '$A$1>4', style: GREEN, priority: 1 }],
		);
		const cf = evaluator(wb);
		expect(cf.at(2, 0)?.style).toEqual(GREEN);
	});

	it('treats formula errors as no match', () => {
		const wb = column([1], [{ type: 'expression', formula: 'BAD(', style: GREEN, priority: 1 }]);
		expect(evaluator(wb).at(0, 0)).toBeUndefined();
	});

	it('merges styles by priority and honours stopIfTrue', () => {
		const both = column(
			[10],
			[
				{
					type: 'cellIs',
					operator: 'greaterThan',
					formulas: ['1'],
					style: { font: { color: { rgb: 'FF0000' } } },
					priority: 2,
				},
				{
					type: 'cellIs',
					operator: 'greaterThan',
					formulas: ['5'],
					style: { font: { color: { rgb: '00FF00' }, bold: true } },
					priority: 1,
				},
			],
		);
		expect(evaluator(both).at(0, 0)?.style?.font).toEqual({ color: { rgb: '00FF00' }, bold: true });
		const stop = column(
			[10],
			[
				{
					type: 'cellIs',
					operator: 'greaterThan',
					formulas: ['5'],
					style: BOLD,
					priority: 1,
					stopIfTrue: true,
				},
				{ type: 'cellIs', operator: 'greaterThan', formulas: ['1'], style: GREEN, priority: 2 },
			],
		);
		expect(evaluator(stop).at(0, 0)?.style).toEqual(BOLD);
		const fill = column(
			[10],
			[
				{ type: 'cellIs', operator: 'greaterThan', formulas: ['5'], style: BOLD, priority: 1 },
				{ type: 'cellIs', operator: 'greaterThan', formulas: ['1'], style: GREEN, priority: 2 },
			],
		);
		expect(evaluator(fill).at(0, 0)?.style).toEqual({ font: { bold: true }, fill: GREEN.fill });
	});

	it('interpolates 2- and 3-colour scales', () => {
		const two = column(
			[0, 50, 100],
			[
				{
					type: 'colorScale',
					thresholds: [{ type: 'min' }, { type: 'max' }],
					colors: [{ rgb: '000000' }, { rgb: 'FFFFFF' }],
					priority: 1,
				},
			],
		);
		const cf = evaluator(two);
		expect(cf.at(0, 0)?.colorScale).toBe('#000000');
		expect(cf.at(1, 0)?.colorScale).toBe('#808080');
		expect(cf.at(2, 0)?.colorScale).toBe('#FFFFFF');
		const three = column(
			[1, 2, 3, 4, 5],
			[
				{
					type: 'colorScale',
					thresholds: [{ type: 'min' }, { type: 'percentile', value: '50' }, { type: 'max' }],
					colors: [{ rgb: 'F8696B' }, { rgb: 'FFEB84' }, { rgb: '63BE7B' }],
					priority: 1,
				},
			],
		);
		const cf3 = evaluator(three);
		expect(cf3.at(0, 0)?.colorScale).toBe('#F8696B');
		expect(cf3.at(2, 0)?.colorScale).toBe('#FFEB84');
		expect(cf3.at(4, 0)?.colorScale).toBe('#63BE7B');
		expect(cf3.at(1, 0)?.colorScale).toBe('#FCAA78');
	});

	it('uses num and percent thresholds', () => {
		const wb = column(
			[0, 10, 20],
			[
				{
					type: 'colorScale',
					thresholds: [
						{ type: 'num', value: '10' },
						{ type: 'percent', value: '100' },
					],
					colors: [{ rgb: '000000' }, { rgb: 'FFFFFF' }],
					priority: 1,
				},
			],
		);
		const cf = evaluator(wb);
		expect(cf.at(0, 0)?.colorScale).toBe('#000000');
		expect(cf.at(1, 0)?.colorScale).toBe('#000000');
		expect(cf.at(2, 0)?.colorScale).toBe('#FFFFFF');
	});

	it('ignores non-numeric cells for scales, bars and icons', () => {
		const wb = column(
			['a', 1, 2],
			[
				{
					type: 'colorScale',
					thresholds: [{ type: 'min' }, { type: 'max' }],
					colors: [{ rgb: '000000' }, { rgb: 'FFFFFF' }],
					priority: 1,
				},
			],
		);
		expect(evaluator(wb).at(0, 0)).toBeUndefined();
	});

	it('computes data bar fractions', () => {
		const wb = column(
			[0, 25, 100, -10],
			[
				{
					type: 'dataBar',
					min: { type: 'num', value: '0' },
					max: { type: 'max' },
					color: { rgb: '638EC6' },
					priority: 1,
				},
			],
		);
		const cf = evaluator(wb);
		expect(cf.at(0, 0)?.dataBar).toEqual({ fraction: 0, color: '#638EC6' });
		expect(cf.at(1, 0)?.dataBar?.fraction).toBe(0.25);
		expect(cf.at(2, 0)?.dataBar?.fraction).toBe(1);
		expect(cf.at(3, 0)?.dataBar).toEqual({ fraction: 0, color: '#638EC6', negative: true });
	});

	it('hides values for showValue=0 bars', () => {
		const wb = column(
			[5],
			[
				{
					type: 'dataBar',
					min: { type: 'min' },
					max: { type: 'max' },
					color: { theme: 4 },
					priority: 1,
					showValue: false,
				},
			],
		);
		const r = evaluator(wb).at(0, 0);
		expect(r?.hideValue).toBe(true);
		expect(r?.dataBar).toEqual({ fraction: 1, color: '#4472C4' });
	});

	it('picks icon indices from percent thresholds', () => {
		const wb = column(
			[0, 32, 33, 66, 67, 100],
			[
				{
					type: 'iconSet',
					iconSet: '3TrafficLights1',
					thresholds: [
						{ type: 'percent', value: '0' },
						{ type: 'percent', value: '33' },
						{ type: 'percent', value: '67' },
					],
					priority: 1,
				},
			],
		);
		const cf = evaluator(wb);
		expect([0, 1, 2, 3, 4, 5].map((r) => cf.at(r, 0)?.icon?.index)).toEqual([0, 0, 1, 1, 2, 2]);
		expect(cf.at(0, 0)?.icon?.set).toBe('3TrafficLights1');
	});

	it('reverses icon order', () => {
		const wb = column(
			[0, 100],
			[
				{
					type: 'iconSet',
					iconSet: '3Arrows',
					thresholds: [
						{ type: 'percent', value: '0' },
						{ type: 'percent', value: '33' },
						{ type: 'percent', value: '67' },
					],
					priority: 1,
					reverse: true,
				},
			],
		);
		const cf = evaluator(wb);
		expect(cf.at(0, 0)?.icon?.index).toBe(2);
		expect(cf.at(1, 0)?.icon?.index).toBe(0);
	});

	it('evaluates top and bottom N, by count and percent', () => {
		const values = [5, 1, 9, 3, 7, 2, 8, 4, 6, 10];
		const top3 = evaluator(column(values, [{ type: 'top10', rank: 3, style: RED, priority: 1 }]));
		expect(values.map((_, i) => top3.at(i, 0) !== undefined)).toEqual(values.map((v) => v >= 8));
		const bottom2 = evaluator(
			column(values, [{ type: 'top10', rank: 2, bottom: true, style: RED, priority: 1 }]),
		);
		expect(values.map((_, i) => bottom2.at(i, 0) !== undefined)).toEqual(values.map((v) => v <= 2));
		const top20pct = evaluator(
			column(values, [{ type: 'top10', rank: 20, percent: true, style: RED, priority: 1 }]),
		);
		expect(values.map((_, i) => top20pct.at(i, 0) !== undefined)).toEqual(
			values.map((v) => v >= 9),
		);
	});

	it('evaluates above and below average', () => {
		const values = [1, 2, 3, 4, 5];
		const above = evaluator(column(values, [{ type: 'aboveAverage', style: RED, priority: 1 }]));
		expect(values.map((_, i) => above.at(i, 0) !== undefined)).toEqual([
			false,
			false,
			false,
			true,
			true,
		]);
		const below = evaluator(
			column(values, [
				{ type: 'aboveAverage', below: true, equalAverage: true, style: RED, priority: 1 },
			]),
		);
		expect(values.map((_, i) => below.at(i, 0) !== undefined)).toEqual([
			true,
			true,
			true,
			false,
			false,
		]);
	});

	it('finds duplicate and unique values', () => {
		const values: CellValue[] = ['a', 'A', 'b', 1, 1, null];
		const dup = evaluator(column(values, [{ type: 'duplicateValues', style: RED, priority: 1 }]));
		expect(values.map((_, i) => dup.at(i, 0) !== undefined)).toEqual([
			true,
			true,
			false,
			true,
			true,
			false,
		]);
		const uniq = evaluator(column(values, [{ type: 'uniqueValues', style: RED, priority: 1 }]));
		expect(values.map((_, i) => uniq.at(i, 0) !== undefined)).toEqual([
			false,
			false,
			true,
			false,
			false,
			false,
		]);
	});

	it('evaluates text rules case-insensitively', () => {
		const values = ['Apple pie', 'banana', 'PIE chart'];
		const has = evaluator(
			column(values, [{ type: 'containsText', text: 'pie', style: RED, priority: 1 }]),
		);
		expect(values.map((_, i) => has.at(i, 0) !== undefined)).toEqual([true, false, true]);
		const not = evaluator(
			column(values, [{ type: 'notContainsText', text: 'pie', style: RED, priority: 1 }]),
		);
		expect(values.map((_, i) => not.at(i, 0) !== undefined)).toEqual([false, true, false]);
		const begins = evaluator(
			column(values, [{ type: 'beginsWith', text: 'pie', style: RED, priority: 1 }]),
		);
		expect(values.map((_, i) => begins.at(i, 0) !== undefined)).toEqual([false, false, true]);
		const ends = evaluator(
			column(values, [{ type: 'endsWith', text: 'PIE', style: RED, priority: 1 }]),
		);
		expect(values.map((_, i) => ends.at(i, 0) !== undefined)).toEqual([true, false, false]);
	});

	it('evaluates blanks and errors', () => {
		const values: CellValue[] = [null, '  ', 'x', { error: '#N/A' }];
		const blanks = evaluator(column(values, [{ type: 'containsBlanks', style: RED, priority: 1 }]));
		expect(values.map((_, i) => blanks.at(i, 0) !== undefined)).toEqual([true, true, false, false]);
		const notBlanks = evaluator(
			column(values, [{ type: 'notContainsBlanks', style: RED, priority: 1 }]),
		);
		expect(values.map((_, i) => notBlanks.at(i, 0) !== undefined)).toEqual([
			false,
			false,
			true,
			true,
		]);
		const errors = evaluator(column(values, [{ type: 'containsErrors', style: RED, priority: 1 }]));
		expect(values.map((_, i) => errors.at(i, 0) !== undefined)).toEqual([
			false,
			false,
			false,
			true,
		]);
		const noErrors = evaluator(
			column(values, [{ type: 'notContainsErrors', style: RED, priority: 1 }]),
		);
		expect(values.map((_, i) => noErrors.at(i, 0) !== undefined)).toEqual([
			true,
			true,
			true,
			false,
		]);
	});

	it('caches statistics until a new evaluator is created', () => {
		const wb = column([1, 2, 3], [{ type: 'aboveAverage', style: RED, priority: 1 }]);
		const cf = evaluator(wb);
		expect(cf.at(2, 0)).toBeDefined();
		putCell(wb.sheets[0]!, 0, 0, { value: 100 });
		expect(cf.at(2, 0)).toBeDefined();
		expect(evaluator(wb).at(2, 0)).toBeUndefined();
	});

	it('returns nothing for a missing sheet', () => {
		const wb = createWorkbook();
		expect(createConditionalFormatEvaluator(wb, 5, () => null).at(0, 0)).toBeUndefined();
	});

	it('blends scale colours between cuts', () => {
		expect(scaleColor(5, [0, 10], ['#000000', '#FFFFFF'])).toBe('#808080');
		expect(scaleColor(5, [5, 5], ['#000000', '#FFFFFF'])).toBe('#000000');
		expect(scaleColor(5, [], [])).toBeUndefined();
	});
});

describe('conditional formats with the real calc engine', () => {
	it('translates relative references with the formula module by default', async () => {
		const { createCalcEngine } = await import('../formula/index.js');
		const wb = column(
			[1, 5, 10],
			[{ type: 'expression', formula: 'A1>4', style: GREEN, priority: 1 }],
		);
		const engine = createCalcEngine(wb);
		const cf = createConditionalFormatEvaluator(wb, 0, (f, at) => engine.evaluate(f, at));
		expect([0, 1, 2].map((r) => cf.at(r, 0) !== undefined)).toEqual([false, true, true]);
	});
});
