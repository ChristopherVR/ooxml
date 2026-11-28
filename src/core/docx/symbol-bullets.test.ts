import { describe, expect, it } from 'vitest';
import { symbolCharToUnicode } from '../text/symbol-fonts';
import { displayListLabel, symbolBullets } from './symbol-bullets';

describe('symbol bullet tables', () => {
	it.each(['Symbol', 'Wingdings'])('agree with the shared %s glyph table', (font) => {
		const table = symbolBullets(font)!;
		expect(table.size).toBeGreaterThan(10);
		for (const [byte, picture] of table) {
			const shared = symbolCharToUnicode(font, String.fromCharCode(0xf000 + byte));
			// Word names Wingdings F0FC U+2713 where the shared (PowerPoint) table has U+2714.
			if (font === 'Wingdings' && byte === 0xfc) expect(shared).toBe('✔');
			else expect([byte, picture]).toEqual([byte, shared]);
		}
	});
});

describe('displayListLabel', () => {
	it('draws a label that is only a soft hyphen as a hyphen (#433)', () => {
		expect(displayListLabel('­', 'Courier New')).toBe('-');
		expect(displayListLabel('a­b', 'Courier New')).toBe('a­b');
	});
});
