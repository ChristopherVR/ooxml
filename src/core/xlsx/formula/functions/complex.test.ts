// Expected strings are the worked examples in Microsoft's function reference.
import { describe, expect, it } from 'vitest';
import { calc, E } from '../test-helpers.js';

describe('complex numbers', () => {
	it('COMPLEX builds the text form, omitting a coefficient of 1', () => {
		expect(calc('COMPLEX(3,4)')).toBe('3+4i');
		expect(calc('COMPLEX(3,4,"j")')).toBe('3+4j');
		expect(calc('COMPLEX(0,1)')).toBe('i');
		expect(calc('COMPLEX(0,-1,"j")')).toBe('-j');
		expect(calc('COMPLEX(5,0)')).toBe('5');
		expect(calc('COMPLEX(3,4,"k")')).toEqual(E.VALUE);
	});

	it('reads the parts and the modulus and argument', () => {
		expect(calc('IMREAL("6-9i")')).toBe(6);
		expect(calc('IMAGINARY("3+4i")')).toBe(4);
		expect(calc('IMAGINARY("-j")')).toBe(-1);
		expect(calc('IMABS("5+12i")')).toBe(13);
		expect(calc('IMARGUMENT("3+4i")')).toBeCloseTo(0.927295218, 8);
		expect(calc('IMARGUMENT(0)')).toEqual(E.DIV0);
		expect(calc('IMCONJUGATE("3+4i")')).toBe('3-4i');
	});

	it('does arithmetic', () => {
		expect(calc('IMSUM("3+4i","5-3i")')).toBe('8+i');
		expect(calc('IMSUB("13+4i","5+3i")')).toBe('8+i');
		expect(calc('IMPRODUCT("3+4i","5-3i")')).toBe('27+11i');
		expect(calc('IMDIV("-238+240i","10+24i")')).toBe('5+12i');
		expect(calc('IMPOWER("2+3i",3)')).toBe('-46+9i');
		expect(calc('IMDIV("1","0")')).toEqual(E.NUM);
		expect(calc('IMSUM("1+i","2+2j")')).toEqual(E.VALUE);
	});

	it('does exponentials, logarithms and roots', () => {
		expect(calc('IMEXP("1+i")')).toBe('1.46869393991589+2.28735528717884i');
		expect(calc('IMLN("3+4i")')).toBe('1.6094379124341+0.927295218001612i');
		expect(calc('IMLOG10("3+4i")')).toBe('0.698970004336019+0.402719196273373i');
		expect(calc('IMLOG2("3+4i")')).toBe('2.32192809488736+1.33780421245098i');
		expect(calc('IMSQRT("1+i")')).toBe('1.09868411346781+0.455089860562227i');
	});

	it('does trigonometry', () => {
		expect(calc('IMSIN("3+4i")')).toBe('3.85373803791938-27.0168132580039i');
		expect(calc('IMCOS("1+i")')).toBe('0.833730025131149-0.988897705762865i');
		// Identities hold to the 15 digits the text form keeps.
		const real = (formula: string) => calc(`IMREAL(${formula})`) as number;
		const imag = (formula: string) => calc(`IMAGINARY(${formula})`) as number;
		const identity = 'IMSUM(IMPOWER(IMSIN("1+2i"),2),IMPOWER(IMCOS("1+2i"),2))';
		expect(real(identity)).toBeCloseTo(1, 10);
		expect(imag(identity)).toBeCloseTo(0, 10);
		const tan = 'IMTAN("1+2i")';
		const quotient = 'IMDIV(IMSIN("1+2i"),IMCOS("1+2i"))';
		expect(real(tan)).toBeCloseTo(real(quotient), 10);
		expect(imag(tan)).toBeCloseTo(imag(quotient), 10);
		expect(real('IMPRODUCT(IMSEC("1+i"),IMCOS("1+i"))')).toBeCloseTo(1, 10);
		expect(real('IMPRODUCT(IMCSCH("1+i"),IMSINH("1+i"))')).toBeCloseTo(1, 10);
		expect(real('IMPRODUCT(IMCOT("1+i"),IMSIN("1+i"))')).toBeCloseTo(real('IMCOS("1+i")'), 10);
	});

	it('rejects text that is not a complex number', () => {
		expect(calc('IMABS("3 + 4i")')).toEqual(E.NUM);
		expect(calc('IMABS("abc")')).toEqual(E.NUM);
		expect(calc('IMSUM(A1:A2)', { A1: '1+i', A2: '2-3i' })).toBe('3-2i');
	});
});
