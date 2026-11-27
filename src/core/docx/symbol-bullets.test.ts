import { describe, expect, it } from 'vitest';
import { symbolCharToUnicode } from '../text/symbol-fonts';
import { symbolBulletTable } from './symbol-bullets';

describe('symbol bullet tables', () => {
	it.each(['Symbol', 'Wingdings'])('agree with the shared %s glyph table', (font) => {
		const table = symbolBulletTable(font)!;
		for (const [byte, picture] of Object.entries(table)) {
			const shared = symbolCharToUnicode(font, String.fromCharCode(0xf000 + Number(byte)));
			// Word names Wingdings F0FC U+2713 where the shared (PowerPoint) table has U+2714.
			if (font === 'Wingdings' && Number(byte) === 0xfc) expect(shared).toBe('✔');
			else expect([byte, picture]).toEqual([byte, shared]);
		}
	});
});
