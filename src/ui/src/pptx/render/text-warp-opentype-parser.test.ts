import { Font, Glyph, Path } from 'opentype.js';
import { describe, expect, it, vi } from 'vitest';

import {
	isOpenTypeFontParserReady,
	parseOpenTypeFont,
	preloadOpenTypeFontParser,
} from './text-warp-opentype-parser';

describe('parseOpenTypeFont module interop', () => {
	it('uses the ESM named parser when available', () => {
		const font = {} as Font;
		const parse = vi.fn(() => font);
		expect(parseOpenTypeFont(new ArrayBuffer(0), { parse })).toBe(font);
		expect(parse).toHaveBeenCalledOnce();
	});

	it('falls back to the CommonJS default object parser', () => {
		const font = {} as Font;
		const parse = vi.fn(() => font);
		expect(parseOpenTypeFont(new ArrayBuffer(0), { default: { parse } })).toBe(font);
		expect(parse).toHaveBeenCalledOnce();
	});

	it('fails clearly when neither module shape exposes a parser', () => {
		expect(() => parseOpenTypeFont(new ArrayBuffer(0), {})).toThrow(
			'opentype.js does not expose a font parser',
		);
	});
});

describe('preloadOpenTypeFontParser', () => {
	it('shares one import, after which parseOpenTypeFont needs no explicit module', async () => {
		const first = preloadOpenTypeFontParser();
		expect(preloadOpenTypeFontParser()).toBe(first);
		await first;
		expect(isOpenTypeFontParserReady()).toBe(true);
		const font = new Font({
			familyName: 'ParserReady',
			styleName: 'Regular',
			unitsPerEm: 1000,
			ascender: 800,
			descender: -200,
			glyphs: [new Glyph({ name: '.notdef', unicode: 0, advanceWidth: 650, path: new Path() })],
		});
		expect(parseOpenTypeFont(font.toArrayBuffer()).unitsPerEm).toBe(1000);
	});
});
