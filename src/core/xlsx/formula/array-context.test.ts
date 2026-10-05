import { describe, expect, it } from 'vitest';
import { needsArrayEvaluation } from './array-context.js';

describe('needsArrayEvaluation', () => {
	it('is false for scalar formulas and array-native aggregation', () => {
		for (const f of [
			'A1+B1',
			'SUM(A1:A5)',
			'VLOOKUP(A1,B1:C9,2,FALSE)',
			'SUMPRODUCT((A1:A5>2)*B1:B5)',
		])
			expect(needsArrayEvaluation(f), f).toBe(false);
	});
	it('is true for multi-valued results and array operations in scalar positions', () => {
		for (const f of ['A1:A5', 'A1:A5*2', 'SUM(A1:A5*B1:B5)', 'ABS(A1:A3)', '{1,2,3}', 'A1#'])
			expect(needsArrayEvaluation(f), f).toBe(true);
	});
	it('is false for an unparsable formula', () => {
		expect(needsArrayEvaluation('SUM((')).toBe(false);
	});
});
