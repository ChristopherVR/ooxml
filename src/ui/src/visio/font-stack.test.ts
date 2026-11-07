import { describe, expect, it } from 'vitest';
import { primaryFontFamily, visioFontStack } from './font-stack';

describe('visioFontStack', () => {
	it('lists the metric-compatible clone after the drawing family', () => {
		const calibri = visioFontStack('Calibri');
		expect(calibri.startsWith('"Calibri", "Carlito"')).toBe(true);
		expect(calibri.endsWith('sans-serif')).toBe(true);
		expect(visioFontStack('Cambria')).toContain('"Caladea"');
		expect(visioFontStack('Georgia')).toContain('"Gelasio"');
	});

	it('ends an unknown family with a generic one', () => {
		expect(visioFontStack('Private Face')).toBe('"Private Face", sans-serif');
	});

	it('cannot be broken out of by a hostile family name', () => {
		expect(visioFontStack('Arial" onload="x')).toMatch(/^"Arial onload=x", /u);
	});
});

describe('primaryFontFamily', () => {
	it('reads back the drawing family from a painted list', () => {
		expect(primaryFontFamily(visioFontStack('Times New Roman'))).toBe('Times New Roman');
		expect(primaryFontFamily(visioFontStack('A, B'))).toBe('A, B');
		expect(primaryFontFamily('Segoe UI, sans-serif')).toBe('Segoe UI');
		expect(primaryFontFamily('')).toBe('');
	});
});
