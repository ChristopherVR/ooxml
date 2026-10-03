import { describe, expect, it } from 'vitest';
import { analyzeVisioFormula, evaluateVisioFormula, VisioFormulaError } from './formula.js';
import type { VisioFormulaValue } from './formula.js';

const resolve = ({ cell }: { cell: string }): VisioFormulaValue => {
	const values: Record<string, VisioFormulaValue> = {
		Width: { value: 3, unit: 'length' },
		Height: { value: 4, unit: 'length' },
		Angle: { value: Math.PI, unit: 'angle' },
		Mask: { value: 12, unit: 'scalar' },
	};
	if (!Object.hasOwn(values, cell)) throw new VisioFormulaError('reference', 'Unknown cell');
	return values[cell]!;
};
describe('bounded numeric ShapeSheet extension', () => {
	it('evaluates logical expressions and retains dependencies from all branches', () => {
		for (const [source, expected] of [
			['AND(1,2,-1)', 1],
			['AND(1,0)', 0],
			['OR(0,-2)', 1],
			['NOT(0)', 1],
			['NOT(Width)', 0],
		] as const)
			expect(evaluateVisioFormula(source, resolve)).toEqual({ value: expected, unit: 'scalar' });
		const analysis = analyzeVisioFormula('IF(AND(Width>0,NOT(Mask)),Height,Width)');
		expect(analysis.unsupportedFunctions).toEqual([]);
		expect(analysis.references.map((ref) => ref.cell).sort()).toEqual(['Height', 'Mask', 'Width']);
		// Eager logical argument evaluation must not hide missing inputs or errors.
		expect(() => evaluateVisioFormula('AND(0,Missing)', resolve)).toThrow(/Unknown cell/);
		expect(() => evaluateVisioFormula('OR(1,1/0)', resolve)).toThrow(/Division by zero/);
	});
	it('evaluates unsigned 16-bit masks without Javascript truncation/coercion', () => {
		for (const [source, expected] of [
			['BITAND(12,6)', 4],
			['BITOR(12,6)', 14],
			['BITXOR(12,6)', 10],
			['BITAND(65535,32768)', 32768],
			['BITXOR(65535,65535)', 0],
		] as const)
			expect(evaluateVisioFormula(source, resolve)).toEqual({ value: expected, unit: 'scalar' });
		for (const source of [
			'BITAND(-1,1)',
			'BITOR(65536,1)',
			'BITXOR(0.5,1)',
			'BITAND(Width,1)',
			'BITAND("12",6)',
			'BITAND(1)',
			'NOT(1,2)',
			'AND()',
		])
			expect(() => evaluateVisioFormula(source, resolve), source).toThrow(VisioFormulaError);
	});
	it('matches Microsoft MODULUS divisor-sign examples and preserves dimensions', () => {
		for (const [source, expected] of [
			['MODULUS(5,1.4)', 0.8],
			['MODULUS(5,-1.4)', -0.6],
			['MODULUS(-5,1.4)', 0.6],
			['MODULUS(-5,-1.4)', -0.8],
		] as const)
			expect(evaluateVisioFormula(source, resolve).value).toBeCloseTo(expected, 12);
		expect(evaluateVisioFormula('MODULUS(Angle,180DEG)', resolve)).toEqual({
			value: 0,
			unit: 'angle',
		});
		expect(evaluateVisioFormula('MODULUS(Width,2in)', resolve)).toEqual({
			value: 1,
			unit: 'length',
		});
		for (const source of [
			'MODULUS(Width,Angle)',
			'MODULUS(1,0)',
			'MODULUS(1)',
			'MODULUS(1e308,1e-308)',
		])
			expect(() => evaluateVisioFormula(source, resolve), source).toThrow(VisioFormulaError);
	});
	it('carries square inches through distance formulas and restores inches only through valid algebra', () => {
		expect(evaluateVisioFormula('Width^2+Height^2', resolve)).toEqual({ value: 25, unit: 'area' });
		expect(evaluateVisioFormula('SQRT(Width^2+Height^2)', resolve)).toEqual({
			value: 5,
			unit: 'length',
		});
		expect(evaluateVisioFormula('Width*Height/Width', resolve)).toEqual({
			value: 4,
			unit: 'length',
		});
		expect(evaluateVisioFormula('(Width^2)/(Height^2)', resolve)).toEqual({
			value: 9 / 16,
			unit: 'scalar',
		});
		for (const source of [
			'Width+Height^2',
			'Width*Angle',
			'Width^3',
			'Width^-1',
			'1/Width',
			'SQRT(Width)',
			'Angle^2',
			'Width^Height',
			'SQRT(-Width^2)',
		])
			expect(() => evaluateVisioFormula(source, resolve), source).toThrow(VisioFormulaError);
	});
	it('retains interpreter and aggregate step budgets for new functions', () => {
		expect(() => evaluateVisioFormula('BITAND(12,6)', resolve, { maxSteps: 2 })).toThrow(
			/evaluation limit/,
		);
		let steps = 0;
		expect(() =>
			evaluateVisioFormula('AND(1,2,3)', resolve, {
				onStep: () => {
					if (++steps > 3) throw new VisioFormulaError('limit', 'aggregate steps');
				},
			}),
		).toThrow(/aggregate steps/);
		expect(() => evaluateVisioFormula('EVALCELL(Width)', resolve)).toThrow(/Unsupported function/);
	});
});
