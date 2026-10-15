import { describe, expect, it } from 'vitest';
import { evaluateVisioFormula, parseVisioFormula, type VisioFormulaValue } from './formula';

const cells: Record<string, VisioFormulaValue> = {
	width: { value: 2, unit: 'length' },
	height: { value: 1, unit: 'length' },
	false: { value: 0, unit: 'scalar' },
	true: { value: 1, unit: 'scalar' },
};
const evaluate = (source: string, options = {}) =>
	evaluateVisioFormula(source, (reference) => cells[reference.cell.toLowerCase()]!, options);

describe('BOUND, percentages and bare inch constants', () => {
	it('keeps a control handle inside its range, as stencil masters ask', () => {
		// Visio's Can and Pyramid: BOUND(value, 0, FALSE, min, max).
		expect(evaluate('BOUND(Height*0.8,0,FALSE,0,Height)')).toEqual({ value: 0.8, unit: 'length' });
		expect(evaluate('BOUND(Width*2,0,FALSE,Width*0,Width*1)').value).toBe(2);
		expect(evaluate('BOUND(-1 in,0,FALSE,Width*0,Width*1)').value).toBe(0);
		// The nearest of several ranges, an ignored range, and a disabled constraint.
		expect(evaluate('BOUND(1.4 in,0,FALSE,0 in,1 in,FALSE,2 in,3 in)').value).toBe(1);
		expect(evaluate('BOUND(1.4 in,0,TRUE,0 in,1 in,FALSE,2 in,3 in)').value).toBe(2);
		expect(evaluate('BOUND(5 in,2,FALSE,0 in,1 in)').value).toBe(5);
		// Excluded ranges (type 1) push a value inside one to its nearest edge, recorded from Visio 16.
		expect(evaluate('BOUND(5 in,1,FALSE,0 in,1 in)').value).toBe(5);
		expect(evaluate('BOUND(1.5,1,FALSE,1,3)').value).toBe(1);
		expect(evaluate('BOUND(2.5,1,FALSE,1,3)').value).toBe(3);
		expect(evaluate('BOUND(2,1,FALSE,1,3)').value).toBe(3);
		expect(evaluate('BOUND(2,1,TRUE,1,3)').value).toBe(2);
		expect(evaluate('BOUND(15,1,FALSE,1,3,FALSE,10,20)').value).toBe(20);
		expect(() => evaluate('BOUND(2,3,FALSE,1,3)')).toThrow(/type/);
		expect(() => evaluate('BOUND(5 in,0,FALSE,0 in)')).toThrow(/argument count/);
	});

	it('reads a percentage literal', () => {
		expect(parseVisioFormula('THEMEVAL("LineColorTrans",0%)').kind).toBe('call');
		expect(evaluate('50%').value).toBe(0.5);
	});

	it('takes a bare constant added to a length as inches only when asked', () => {
		expect(() => evaluate('Width*0.5+0.0015')).toThrow(/Incompatible units/);
		expect(evaluate('Width*0.5+0.0015', { bareLengths: true })).toEqual({
			value: 1.0015,
			unit: 'length',
		});
		expect(evaluate('0.25-Width', { bareLengths: true }).value).toBe(-1.75);
		// A bare cell value is not a literal: it keeps its own unit.
		expect(() => evaluate('Width+TRUE', { bareLengths: true })).toThrow(/Incompatible units/);
	});
});
