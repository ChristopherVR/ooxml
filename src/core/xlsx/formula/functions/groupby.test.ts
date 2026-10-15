import { describe, expect, it } from 'vitest';
import { calcArray, E } from '../test-helpers';

const data = {
	A1: 'Region',
	B1: 'Sales',
	C1: 'Kind',
	A2: 'West',
	B2: 10,
	C2: 'x',
	A3: 'East',
	B3: 5,
	C3: 'y',
	A4: 'West',
	B4: 20,
	C4: 'y',
	A5: 'East',
	B5: 7,
	C5: 'x',
};

describe('GROUPBY', () => {
	it('groups, sorts ascending and adds a grand total', () => {
		expect(calcArray('GROUPBY(A1:A5,B1:B5,SUM)', data, 4, 2)).toEqual([
			['Region', 'Sales'],
			['East', 12],
			['West', 30],
			['Total', 42],
		]);
	});
	it('sorts descending, drops the total, and accepts a LAMBDA', () => {
		expect(calcArray('GROUPBY(A2:A5,B2:B5,LAMBDA(v,MAX(v)),0,0,-2)', data, 2, 2)).toEqual([
			['West', 20],
			['East', 7],
		]);
	});
	it('supports PERCENTOF, filter_array and several row fields', () => {
		expect(calcArray('GROUPBY(A1:A5,B1:B5,PERCENTOF,,0)', data, 3, 2)).toEqual([
			['Region', 'Sales'],
			['East', 12 / 42],
			['West', 30 / 42],
		]);
		expect(
			calcArray('GROUPBY(HSTACK(A2:A5,C2:C5),B2:B5,SUM,0,0,,A2:A5="West")', data, 2, 3),
		).toEqual([
			['West', 'x', 10],
			['West', 'y', 20],
		]);
	});
	it('rejects unsupported total depths', () => {
		expect(calcArray('GROUPBY(A2:A5,B2:B5,SUM,0,2)', data, 1, 1)).toEqual([[E.VALUE]]);
	});
});

describe('PIVOTBY', () => {
	it('cross-tabulates with totals', () => {
		expect(calcArray('PIVOTBY(A2:A5,C2:C5,B2:B5,SUM,0)', data, 4, 4)).toEqual([
			['', 'x', 'y', 'Total'],
			['East', 7, 5, 12],
			['West', 10, 20, 30],
			['Total', 17, 25, 42],
		]);
	});
});
