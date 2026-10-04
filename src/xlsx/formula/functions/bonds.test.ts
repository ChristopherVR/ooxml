// Expected values are the worked examples in Microsoft's function reference for each function.
import { describe, expect, it } from 'vitest';
import { calc, E } from '../test-helpers.js';

const close = (formula: string, expected: number, digits: number) =>
	expect(calc(formula) as number).toBeCloseTo(expected, digits);

describe('coupon schedule', () => {
	const args = 'DATE(2011,1,25),DATE(2011,11,15),2,1';
	it('finds the surrounding coupon dates and counts', () => {
		expect(calc(`COUPPCD(${args})`)).toBe(calc('DATE(2010,11,15)'));
		expect(calc(`COUPNCD(${args})`)).toBe(calc('DATE(2011,5,15)'));
		expect(calc(`COUPNUM(${args})`)).toBe(2);
		expect(calc(`COUPDAYBS(${args})`)).toBe(71);
		expect(calc(`COUPDAYS(${args})`)).toBe(181);
		expect(calc(`COUPDAYSNC(${args})`)).toBe(110);
	});

	it('keeps month-end coupon dates when maturity is a month end', () => {
		expect(calc('COUPPCD(DATE(2011,3,15),DATE(2011,8,31),2,0)')).toBe(calc('DATE(2011,2,28)'));
	});

	it('rejects bad frequency, basis and date order', () => {
		expect(calc('COUPNUM(DATE(2011,1,25),DATE(2011,11,15),3,1)')).toEqual(E.NUM);
		expect(calc('COUPNUM(DATE(2011,1,25),DATE(2011,11,15),2,5)')).toEqual(E.NUM);
		expect(calc('COUPNUM(DATE(2011,11,15),DATE(2011,1,25),2,1)')).toEqual(E.NUM);
	});
});

describe('price, yield and duration', () => {
	it('PRICE', () => {
		close('PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)', 94.63436162, 6);
	});
	it('YIELD inverts PRICE', () => {
		close('YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)', 0.065, 6);
		const price = calc('PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.081,100,2,1)') as number;
		close(`YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,${price},100,2,1)`, 0.081, 9);
	});
	it('a security in its last coupon period uses the single-coupon formula', () => {
		const price = calc('PRICE(DATE(2008,2,15),DATE(2008,6,15),0.05,0.06,100,2,0)') as number;
		expect(price).toBeGreaterThan(99);
		expect(price).toBeLessThan(101);
		close(`YIELD(DATE(2008,2,15),DATE(2008,6,15),0.05,${price},100,2,0)`, 0.06, 9);
	});
	it('DURATION and MDURATION', () => {
		close('DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)', 10.9191453, 6);
		close('MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)', 5.73567, 5);
	});
	it('validates negative rates and prices', () => {
		expect(calc('PRICE(DATE(2008,2,15),DATE(2017,11,15),-0.1,0.065,100,2,0)')).toEqual(E.NUM);
		expect(calc('YIELD(DATE(2008,2,15),DATE(2017,11,15),0.05,0,100,2,0)')).toEqual(E.NUM);
	});
});

describe('accrued interest', () => {
	it('ACCRINT from issue and from the first interest date', () => {
		close('ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0)', 16.66667, 4);
		// Settlement after the first interest date: the later periods add up.
		const full = calc(
			'ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,11,30),0.1,1000,2,0)',
		) as number;
		const fromFirst = calc(
			'ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,11,30),0.1,1000,2,0,FALSE)',
		) as number;
		expect(full).toBeGreaterThan(fromFirst);
		expect(fromFirst).toBeCloseTo(1000 * 0.1 * (90 / 360), 6);
	});
	it('ACCRINTM', () => {
		close('ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,3)', 20.54795, 5);
	});
});

describe('discount securities and Treasury bills', () => {
	it('DISC, PRICEDISC and YIELDDISC', () => {
		close('DISC(DATE(2007,1,25),DATE(2007,6,15),97.975,100,1)', 0.052420213, 8);
		close('PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)', 99.79583, 5);
		close('YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)', 0.052823, 6);
	});
	it('INTRATE and RECEIVED', () => {
		close('INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2)', 0.05768, 5);
		close('RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575,2)', 1014584.654, 2);
	});
	it('PRICEMAT and YIELDMAT', () => {
		close(
			'PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)',
			99.98449888,
			7,
		);
		close(
			'YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)',
			0.060954,
			6,
		);
	});
	it('TBILLPRICE, TBILLYIELD and TBILLEQ', () => {
		close('TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)', 98.45, 6);
		close('TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)', 0.09141696, 7);
		close('TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)', 0.094151, 6);
		// Over half a year the equivalent yield uses the quadratic form.
		const long = calc('TBILLEQ(DATE(2008,1,1),DATE(2008,12,1),0.05)') as number;
		expect(long).toBeGreaterThan(0.05);
		expect(long).toBeLessThan(0.06);
	});
	it('rejects a bill longer than a year and non-positive amounts', () => {
		expect(calc('TBILLPRICE(DATE(2008,1,1),DATE(2009,6,1),0.05)')).toEqual(E.NUM);
		expect(calc('DISC(DATE(2007,1,25),DATE(2007,6,15),0,100,1)')).toEqual(E.NUM);
		expect(calc('PRICEDISC(DATE(2008,3,1),DATE(2008,2,16),0.05,100,2)')).toEqual(E.NUM);
	});
});
