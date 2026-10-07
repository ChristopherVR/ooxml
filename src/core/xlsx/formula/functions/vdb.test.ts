// The first three values are the examples in Microsoft's VDB reference.
import { describe, expect, it } from 'vitest';
import { calc, E } from '../test-helpers';

const close = (formula: string, expected: number, digits = 5) =>
	expect(calc(formula) as number).toBeCloseTo(expected, digits);

describe('VDB', () => {
	it('matches the documented examples', () => {
		close('VDB(2400,300,10*365,0,1)', 1.31507, 5);
		close('VDB(2400,300,10*12,0,1)', 40, 8);
		close('VDB(2400,300,10*12,6,18)', 396.31, 2);
		// Factor 3 never switches here, so it is the closed form 2400 * (0.975^6 - 0.975^18).
		close('VDB(2400,300,10*12,6,18,3)', 2400 * (0.975 ** 6 - 0.975 ** 18), 8);
	});

	it('over the whole life depreciates down to salvage', () => {
		close('VDB(2400,300,10,0,10)', 2100, 8);
		close('VDB(2400,300,10,0,10,2,TRUE)', 2100 - 0, 0);
	});

	it('equals DDB for a single period when no switch happens', () => {
		close('VDB(2400,300,10,2,3,2,TRUE)', calc('DDB(2400,300,10,3)') as number, 8);
	});

	it('weights partial periods by the part covered', () => {
		const whole = calc('VDB(2400,300,10,0,1)') as number;
		close('VDB(2400,300,10,0,0.5)', whole / 2, 8);
	});

	it('rejects impossible arguments', () => {
		expect(calc('VDB(2400,300,10,5,3)')).toEqual(E.NUM);
		expect(calc('VDB(2400,300,10,0,11)')).toEqual(E.NUM);
		expect(calc('VDB(-1,300,10,0,1)')).toEqual(E.NUM);
		expect(calc('VDB(2400,300,0,0,1)')).toEqual(E.NUM);
	});
});
