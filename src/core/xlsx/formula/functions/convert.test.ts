// Values follow Microsoft's CONVERT examples and the unit definitions they cite.
import { describe, expect, it } from 'vitest';
import { calc, E } from '../test-helpers.js';

const close = (formula: string, expected: number, digits = 6) =>
	expect(calc(formula) as number).toBeCloseTo(expected, digits);

describe('CONVERT', () => {
	it('converts within a dimension', () => {
		close('CONVERT(1,"lbm","kg")', 0.453592, 6);
		close('CONVERT(1,"mi","km")', 1.609344, 9);
		close('CONVERT(1,"hr","mn")', 60, 9);
		close('CONVERT(1,"gal","l")', 3.785411784, 9);
		close('CONVERT(1,"atm","Pa")', 101325, 6);
		close('CONVERT(1,"mph","m/s")', 0.44704, 9);
		close('CONVERT(CONVERT(100,"ft","m"),"ft","m")', 9.290304, 6);
	});

	it('handles temperature scales with offsets', () => {
		close('CONVERT(68,"F","C")', 20, 9);
		close('CONVERT(0,"C","K")', 273.15, 9);
		close('CONVERT(212,"fah","cel")', 100, 9);
		close('CONVERT(100,"C","F")', 212, 9);
		close('CONVERT(491.67,"Rank","K")', 273.15, 9);
	});

	it('applies metric and binary prefixes, squared for area and cubed for volume', () => {
		close('CONVERT(1,"km","m")', 1000, 9);
		close('CONVERT(1,"mg","g")', 0.001, 12);
		close('CONVERT(1,"km2","m2")', 1e6, 3);
		close('CONVERT(1,"cm3","l")', 0.001, 12);
		close('CONVERT(1,"kbyte","byte")', 1000, 9);
		close('CONVERT(1,"kibyte","byte")', 1024, 9);
		close('CONVERT(1,"byte","bit")', 8, 9);
	});

	it('is #N/A for unknown units, mixed dimensions and a prefix on a unit that cannot take one', () => {
		expect(calc('CONVERT(2.5,"ft","sec")')).toEqual(E.NA);
		expect(calc('CONVERT(1,"parsecs","m")')).toEqual(E.NA);
		expect(calc('CONVERT(1,"kft","m")')).toEqual(E.NA);
		expect(calc('CONVERT(1,"F","m")')).toEqual(E.NA);
		expect(calc('CONVERT(1,"M","m")')).toEqual(E.NA);
	});

	it('unit names are case sensitive', () => {
		close('CONVERT(1,"l","L")', 1, 12);
		expect(calc('CONVERT(1,"KG","g")')).toEqual(E.NA);
	});
});
