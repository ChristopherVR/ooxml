// Records and criteria in the layout of Microsoft's database-function examples; the expected
// values are worked out by hand from the six rows below.
import { describe, expect, it } from 'vitest';
import { calc, E } from '../test-helpers.js';

const rows = [
	['Tree', 'Height', 'Age', 'Yield', 'Profit'],
	['Apple', 18, 20, 14, 105],
	['Pear', 12, 12, 10, 96],
	['Cherry', 13, 14, 9, 105],
	['Apple', 14, 15, 10, 75],
	['Pear', 9, 8, 8, 76.8],
	['Apple', 8, 9, 6, 45],
] as const;
const cols = 'ABCDE';
const orchard: Record<string, string | number> = {};
rows.forEach((r, i) => r.forEach((v, c) => (orchard[`${cols[c]}${i + 10}`] = v)));
// DB is A10:E16.
const DB = 'A10:E16';
const withCriteria = (criteria: Record<string, string | number>) => ({ ...orchard, ...criteria });

const apple = withCriteria({ A1: 'Tree', A2: 'Apple' });
const appleOrPear = withCriteria({ A1: 'Tree', A2: 'Apple', A3: 'Pear' });
const tallApple = withCriteria({ A1: 'Tree', B1: 'Height', A2: 'Apple', B2: '>10' });

describe('database functions', () => {
	it('DSUM adds the matching rows, by label or by column number', () => {
		expect(calc(`DSUM(${DB},"Profit",A1:A2)`, apple)).toBe(225);
		expect(calc(`DSUM(${DB},5,A1:A2)`, apple)).toBe(225);
		expect(calc(`DSUM(${DB},"Profit",A1:A3)`, appleOrPear)).toBeCloseTo(397.8, 10);
	});

	it('criteria on one row are ANDed, rows are ORed', () => {
		expect(calc(`DCOUNT(${DB},"Age",A1:B2)`, tallApple)).toBe(2);
		expect(calc(`DAVERAGE(${DB},"Yield",A1:B2)`, tallApple)).toBe(12);
	});

	it('DCOUNTA counts non-blank cells, DCOUNT only numbers', () => {
		expect(calc(`DCOUNTA(${DB},"Tree",A1:A2)`, apple)).toBe(3);
		expect(calc(`DCOUNT(${DB},"Tree",A1:A2)`, apple)).toBe(0);
	});

	it('DMAX, DMIN and DPRODUCT', () => {
		expect(calc(`DMAX(${DB},"Profit",A1:A3)`, appleOrPear)).toBe(105);
		expect(calc(`DMIN(${DB},"Profit",A1:A3)`, appleOrPear)).toBe(45);
		expect(calc(`DPRODUCT(${DB},"Yield",A1:A2)`, apple)).toBe(840);
	});

	it('DGET needs exactly one match', () => {
		const one = withCriteria({ A1: 'Tree', B1: 'Height', A2: 'Pear', B2: '<10' });
		expect(calc(`DGET(${DB},"Yield",A1:B2)`, one)).toBe(8);
		expect(calc(`DGET(${DB},"Yield",A1:A2)`, withCriteria({ A1: 'Tree', A2: 'Plum' }))).toEqual(
			E.VALUE,
		);
		expect(calc(`DGET(${DB},"Yield",A1:A2)`, apple)).toEqual(E.NUM);
	});

	it('standard deviation and variance, sample and population', () => {
		expect(calc(`DSTDEV(${DB},"Yield",A1:A2)`, apple)).toBeCloseTo(4, 10);
		expect(calc(`DVAR(${DB},"Yield",A1:A2)`, apple)).toBeCloseTo(16, 10);
		expect(calc(`DVARP(${DB},"Yield",A1:A2)`, apple)).toBeCloseTo(32 / 3, 10);
		expect(calc(`DSTDEVP(${DB},"Yield",A1:A2)`, apple)).toBeCloseTo(Math.sqrt(32 / 3), 10);
	});

	it('a blank criteria cell matches every row', () => {
		const blank = withCriteria({ A1: 'Tree', A2: '', B1: 'Height', B2: '>=12' });
		expect(calc(`DCOUNT(${DB},"Height",A1:B2)`, blank)).toBe(4);
	});

	it('reports a bad field and an unknown criteria column', () => {
		expect(calc(`DSUM(${DB},"Nope",A1:A2)`, apple)).toEqual(E.VALUE);
		expect(calc(`DSUM(${DB},9,A1:A2)`, apple)).toEqual(E.VALUE);
		expect(calc(`DSUM(${DB},"Profit",A1:A2)`, withCriteria({ A1: 'Nope', A2: 'x' }))).toBe(0);
	});

	it('DAVERAGE of no matching rows is #DIV/0!', () => {
		expect(calc(`DAVERAGE(${DB},"Yield",A1:A2)`, withCriteria({ A1: 'Tree', A2: 'Plum' }))).toEqual(
			E.DIV0,
		);
	});
});
