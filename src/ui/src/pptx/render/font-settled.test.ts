import { describe, expect, it } from 'vitest';
import { firstFontFamily } from './font-settled';

describe('firstFontFamily', () => {
	it('reads the first family of a font shorthand', () => {
		expect(firstFontFamily('italic 700 40px "Open Sans", Arial')).toBe('Open Sans');
		expect(firstFontFamily("12pt/1.5 'Segoe UI', sans-serif")).toBe('Segoe UI');
		expect(firstFontFamily('16px / 20px Calibri')).toBe('Calibri');
		expect(firstFontFamily('100% Georgia, serif')).toBe('Georgia');
		expect(firstFontFamily('bold Calibri')).toBeUndefined();
		expect(firstFontFamily('12px/1.5')).toBeUndefined();
	});

	it('stays fast on long crafted input (code scanning: polynomial pattern)', () => {
		const started = performance.now();
		firstFontFamily(`0%/${'0%/!'.repeat(20000)}`);
		firstFontFamily(`0%	${'		'.repeat(20000)}`);
		expect(performance.now() - started).toBeLessThan(500);
	});
});
