import { describe, expect, it } from 'vitest';
import { calc, E } from '../test-helpers';

const near = (formula: string, expected: number, digits = 6) =>
	expect(calc(formula) as number).toBeCloseTo(expected, digits);

describe('odd-period securities (Microsoft examples)', () => {
	it('ODDFPRICE', () => {
		near(
			'ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)',
			113.597717,
		);
	});
	it('ODDFYIELD', () => {
		near(
			'ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)',
			0.0772,
			4,
		);
	});
	it('ODDLPRICE', () => {
		near(
			'ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)',
			99.87829,
			5,
		);
	});
	it('ODDLYIELD', () => {
		near(
			'ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)',
			0.0451923,
			6,
		);
	});
	it('rejects out-of-order dates', () => {
		expect(
			calc(
				'ODDFPRICE(DATE(2009,4,1),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.07,0.06,100,2,1)',
			),
		).toEqual(E.NUM);
	});
});

describe('AMORDEGRC', () => {
	it('matches the Microsoft example', () => {
		expect(calc('AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)')).toBe(776);
	});
});

describe('ISOMITTED', () => {
	it('detects omitted LAMBDA arguments', () => {
		expect(calc('LAMBDA(x,y,IF(ISOMITTED(y),"omitted","given"))(1)')).toBe('omitted');
		expect(calc('LAMBDA(x,y,IF(ISOMITTED(y),"omitted","given"))(1,2)')).toBe('given');
		expect(calc('LAMBDA(x,y,IF(ISOMITTED(y),"omitted","given"))(1,)')).toBe('omitted');
	});
});

describe('PERCENTOF and TRIMRANGE', () => {
	it('PERCENTOF divides sums', () => {
		expect(calc('PERCENTOF(A1:A2,A1:A4)', { A1: 1, A2: 2, A3: 3, A4: 4 })).toBe(0.3);
		expect(calc('PERCENTOF(A1,B1)', { A1: 1, B1: 0 })).toEqual(E.DIV0);
	});
	it('TRIMRANGE drops blank edges', () => {
		expect(calc('ROWS(TRIMRANGE(A1:C5))', { B2: 1, B3: 2 })).toBe(2);
		expect(calc('COLUMNS(TRIMRANGE(A1:C5))', { B2: 1, B3: 2 })).toBe(1);
		expect(calc('ROWS(TRIMRANGE(A1:C5,1))', { B2: 1, B3: 2 })).toBe(4);
	});
});

describe('BAHTTEXT', () => {
	it('spells numbers in Thai', () => {
		expect(calc('BAHTTEXT(1234)')).toBe('หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน');
		expect(calc('BAHTTEXT(11)')).toBe('สิบเอ็ดบาทถ้วน');
		expect(calc('BAHTTEXT(21.5)')).toBe('ยี่สิบเอ็ดบาทห้าสิบสตางค์');
		expect(calc('BAHTTEXT(0)')).toBe('ศูนย์บาทถ้วน');
		expect(calc('BAHTTEXT(1000000)')).toBe('หนึ่งล้านบาทถ้วน');
		expect(calc('BAHTTEXT(0.25)')).toBe('ยี่สิบห้าสตางค์');
	});
});
