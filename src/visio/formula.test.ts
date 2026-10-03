import { describe, expect, it } from 'vitest';
import {
	analyzeVisioFormula,
	evaluateVisioFormula,
	parseVisioFormula,
	VisioFormulaError,
	visioFormulaCachedValue,
} from './formula.js';
import type { VisioFormulaReference, VisioFormulaValue } from './formula.js';

const length = (value: number): VisioFormulaValue => ({ value, unit: 'length' });
const scalar = (value: number): VisioFormulaValue => ({ value, unit: 'scalar' });
const resolve = (ref: VisioFormulaReference): VisioFormulaValue => {
	const values: Record<string, VisioFormulaValue> = {
		Width: length(6),
		Height: length(4),
		'Geometry1.X1': length(3),
		'5!Width': length(8),
	};
	const key = ref.shapeId ? `${ref.shapeId}!${ref.cell}` : ref.cell;
	if (!values[key]) throw new VisioFormulaError('reference', `Unknown cell ${key}`);
	return values[key];
};
describe('bounded Visio ShapeSheet formulas', () => {
	it('evaluates common dimensional geometry and same-sheet references', () => {
		expect(evaluateVisioFormula('GUARD(Width*0.5)', resolve)).toEqual(length(3));
		expect(evaluateVisioFormula('Geometry1.X1+Height/2', resolve)).toEqual(length(5));
		expect(evaluateVisioFormula('Sheet.5!Width/Width', resolve)).toEqual(scalar(8 / 6));
		expect(evaluateVisioFormula('IF(Width>0,Height/2,Width)', resolve)).toEqual(length(2));
	});
	it('records references from every IF branch without eagerly evaluating the unused branch', () => {
		const analysis = analyzeVisioFormula('IF(Width>0,Height,Sheet.5!Width)');
		expect(analysis.references).toHaveLength(3);
		expect(analysis.references).toContainEqual({ shapeId: '5', cell: 'Width' });
		expect(evaluateVisioFormula('IF(1,Width,Missing)', resolve)).toEqual(length(6));
	});
	it('uses internal cache units and converts explicit literals', () => {
		expect(visioFormulaCachedValue('1', 'MM')).toEqual(length(1));
		expect(visioFormulaCachedValue('1', 'DP')).toEqual(length(1));
		expect(visioFormulaCachedValue('1', 'DT')).toEqual(length(1));
		expect(visioFormulaCachedValue('1', 'DA')).toEqual({ value: 1, unit: 'angle' });
		expect(evaluateVisioFormula('25.4 mm + 1 in', resolve)).toEqual(length(2));
		expect(evaluateVisioFormula('180 deg', resolve)).toEqual({ value: Math.PI, unit: 'angle' });
	});
	it('implements documented ATAN2 y,x and SIGN fuzz', () => {
		expect(evaluateVisioFormula('ATAN2(1,0)', resolve)).toEqual({
			value: Math.PI / 2,
			unit: 'angle',
		});
		expect(evaluateVisioFormula('SIGN(0.0000000005)', resolve)).toEqual(scalar(0));
		expect(evaluateVisioFormula('SIGN(-0.000000002)', resolve)).toEqual(scalar(-1));
		expect(evaluateVisioFormula('SIGN(0.01,0.1)', resolve)).toEqual(scalar(0));
	});
	it('honors precedence, exponent association, unary signs, and scientific notation', () => {
		expect(evaluateVisioFormula('=2+3*4', resolve)).toEqual(scalar(14));
		expect(evaluateVisioFormula('2^3^2', resolve)).toEqual(scalar(512));
		expect(evaluateVisioFormula('-2^2', resolve)).toEqual(scalar(-4));
		expect(evaluateVisioFormula('1e-3+2E-3', resolve)).toEqual(scalar(0.003));
	});
	it('analyzes unsupported static calls and flags redirection/dynamic references', () => {
		const analysis = analyzeVisioFormula('GUARD(ROUND(Width,Height))');
		expect(analysis.guarded).toBe(true);
		expect(analysis.unsupportedFunctions).toEqual(['ROUND']);
		expect(analysis.references).toHaveLength(2);
		expect(analysis.dynamic).toBe(false);
		expect(analyzeVisioFormula('SETATREF(Width)').dynamic).toBe(true);
		expect(analyzeVisioFormula('INDIRECT(Width)').dynamic).toBe(true);
		expect(analyzeVisioFormula('UNKNOWN(Width)').dynamic).toBe(true);
		expect(() => evaluateVisioFormula('ROUND(Width,Height)', resolve)).toThrow(
			/Unsupported function/,
		);
	});
	it('rejects unknown references rather than using stale cached inputs', () => {
		expect(() => evaluateVisioFormula('Missing+1', resolve)).toThrow(/Unknown cell/);
	});
	it('records implicit theme selectors without pretending theme evaluation is supported', () => {
		const themed = analyzeVisioFormula('THEMEGUARD(THEMEVAL("LineColor"))');
		expect(themed.dynamic).toBe(false);
		expect(themed.guarded).toBe(false);
		expect(themed.references).toContainEqual({ cell: 'QuickStyleLineColor' });
		expect(themed.references).toContainEqual({ cell: 'ColorSchemeIndex' });
		expect(analyzeVisioFormula('ISTHEMED()').references).toEqual([{ cell: 'ColorSchemeIndex' }]);
		expect(analyzeVisioFormula('THEMEVAL(Width)').references).toContainEqual({ cell: 'Width' });
		expect(() => evaluateVisioFormula('THEMEVAL()', resolve)).toThrow(/Unsupported function/);
		expect(analyzeVisioFormula('BITAND(Width,2)').references).toEqual([{ cell: 'Width' }]);
	});
	it('rejects dimensions, division by zero, nonfinite results and invalid arity', () => {
		for (const formula of [
			'Width+1 rad',
			'Width*Height',
			'1/0',
			'1e309',
			'SQRT(-1)',
			'IF(1,2)',
			'GUARD(1,2)',
			'SIN(Width)',
		]) {
			expect(() => evaluateVisioFormula(formula, resolve), formula).toThrow(VisioFormulaError);
		}
		expect(() => visioFormulaCachedValue('NaN')).toThrow(VisioFormulaError);
		expect(() => visioFormulaCachedValue('1', 'unsupported')).toThrow(/unit/);
	});
	it('rejects script fragments and unsupported syntax while analyzing literal strings', () => {
		for (const formula of [
			'globalThis.process.exit()',
			'Width;Height',
			'new Function(1)',
			'Width[0]',
			'Width&&Height',
		]) {
			expect(() => parseVisioFormula(formula), formula).toThrow(VisioFormulaError);
		}
		expect(analyzeVisioFormula('RGB(255,0,0)').dynamic).toBe(false);
		expect(analyzeVisioFormula('IF(0,"Sheet.5!Width",Width)').references).toEqual([
			{ cell: 'Width' },
		]);
		expect(() => evaluateVisioFormula('"Width"', resolve)).toThrow(/String results/);
		expect(() => parseVisioFormula('"unterminated')).toThrow(/Unterminated/);
		expect(parseVisioFormula('"quoted ""value"""')).toEqual({
			kind: 'string',
			value: 'quoted "value"',
		});
	});
	it('enforces length, node, recursion and evaluation step budgets', () => {
		expect(() => parseVisioFormula('1'.repeat(8193))).toThrow(/length limit/);
		expect(() => parseVisioFormula('1+2+3', { maxNodes: 4 })).toThrow(/AST limit/);
		expect(() => parseVisioFormula('('.repeat(70) + '1' + ')'.repeat(70))).toThrow(/depth limit/);
		expect(() => evaluateVisioFormula('1+2', resolve, { maxSteps: 2 })).toThrow(/evaluation limit/);
		expect(() => parseVisioFormula('1', { maxNodes: Infinity })).toThrow(/limit/);
		const deep = Array(80).fill('1').join('+');
		expect(() => analyzeVisioFormula(deep)).toThrow(/analysis limit/);
	});
});
