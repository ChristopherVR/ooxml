import { Font, Glyph, Path } from 'opentype.js';
import type { PptxEmbeddedFont } from 'ooxml-core/pptx';
import { beforeAll, describe, expect, it } from 'vitest';

import { preloadOpenTypeFontParser } from './text-warp-opentype-parser';
import {
	createGlyphOutlineLookup,
	extractGlyphOutlineCommands,
	GlyphOutlineFontCache,
} from './text-warp-outline-font-cache';

/**
 * A minimal, self-contained TrueType-shaped font (`.notdef` plus a single
 * 'A' glyph drawn as a rectangle), built with `opentype.js`'s own object
 * model so this test needs no binary fixture on disk.
 */
function buildTestFont(): Font {
	const notdefGlyph = new Glyph({
		name: '.notdef',
		unicode: 0,
		advanceWidth: 650,
		path: new Path(),
	});
	const aPath = new Path();
	aPath.moveTo(100, 0);
	aPath.lineTo(100, 700);
	aPath.lineTo(300, 700);
	aPath.lineTo(300, 0);
	aPath.close();
	const aGlyph = new Glyph({ name: 'A', unicode: 65, advanceWidth: 650, path: aPath });
	return new Font({
		familyName: 'WarpOutlineTestFont',
		styleName: 'Regular',
		unitsPerEm: 1000,
		ascender: 800,
		descender: -200,
		glyphs: [notdefGlyph, aGlyph],
	});
}

function testFontBytes(): Uint8Array {
	return new Uint8Array(buildTestFont().toArrayBuffer());
}

// These cases cover parsing itself; text-warp-outline-font-cache-lazy.test.ts
// covers what happens before the parser has loaded.
beforeAll(() => preloadOpenTypeFontParser());

describe('extractGlyphOutlineCommands', () => {
	const font = buildTestFont();

	it('returns undefined for an empty char', () => {
		expect(extractGlyphOutlineCommands(font, '', 0, 0, 24)).toBeUndefined();
	});

	it('returns [] for whitespace (a legitimate "nothing to draw" result)', () => {
		expect(extractGlyphOutlineCommands(font, ' ', 0, 0, 24)).toStrictEqual([]);
	});

	it('extracts a real outline for a shaped character, positioned and scaled', () => {
		const commands = extractGlyphOutlineCommands(font, 'A', 50, 100, 1000);
		expect(commands).toBeDefined();
		expect(commands!.length).toBeGreaterThan(0);
		expect(commands![0].type).toBe('M');
		// unitsPerEm=1000, fontSize=1000 => scale=1: the glyph's first point
		// (100, 0) in font units lands at (50+100, 100-0) = (150, 100) in the
		// (x, y)-positioned, y-down SVG space extractGlyphOutlineCommands uses.
		const first = commands![0] as { x: number; y: number };
		expect(first.x).toBeCloseTo(150, 5);
		expect(first.y).toBeCloseTo(100, 5);
	});

	it('returns undefined for a character the font cannot shape', () => {
		// unicode 0x2603 (snowman) is not in this minimal test font.
		expect(extractGlyphOutlineCommands(font, '☃', 0, 0, 24)).toBeUndefined();
	});
});

describe('the GlyphOutlineFontCache class', () => {
	it('parses embedded-font bytes synchronously and caches by (family, bold, italic)', () => {
		const cache = new GlyphOutlineFontCache();
		const embedded: PptxEmbeddedFont = {
			name: 'WarpOutlineTestFont',
			dataUrl: '',
			rawFontData: testFontBytes(),
			bold: false,
			italic: false,
		};
		expect(cache.get('WarpOutlineTestFont', false, false)).toBeUndefined();
		cache.registerEmbeddedFonts([embedded]);
		const parsed = cache.get('WarpOutlineTestFont', false, false);
		expect(parsed).toBeDefined();
		expect(parsed!.hasChar('A')).toBeTruthy();
		// A different style variant was never registered.
		expect(cache.get('WarpOutlineTestFont', true, false)).toBeUndefined();
	});

	it('is idempotent: calling registerEmbeddedFonts again does not re-parse or throw', () => {
		const cache = new GlyphOutlineFontCache();
		const embedded: PptxEmbeddedFont = {
			name: 'WarpOutlineTestFont',
			dataUrl: '',
			rawFontData: testFontBytes(),
		};
		cache.registerEmbeddedFonts([embedded]);
		cache.registerEmbeddedFonts([embedded]);
		expect(cache.get('WarpOutlineTestFont', false, false)).toBeDefined();
		expect(cache.hasAttempted('WarpOutlineTestFont', false, false)).toBeTruthy();
	});

	it('leaves an unparseable font unset rather than throwing', () => {
		const cache = new GlyphOutlineFontCache();
		const bogus: PptxEmbeddedFont = {
			name: 'Bogus',
			dataUrl: '',
			rawFontData: new Uint8Array([1, 2, 3, 4, 5]),
		};
		expect(() => cache.registerEmbeddedFonts([bogus])).not.toThrow();
		expect(cache.get('Bogus', false, false)).toBeUndefined();
		expect(cache.hasAttempted('Bogus', false, false)).toBeTruthy();
	});

	it('matches a CSS font-family fallback LIST against the plain registered name', () => {
		// `getSubstituteFontFamily` (ooxml-core/pptx) returns a full CSS
		// font-family value like `'"WarpOutlineTestFont", "Fallback", sans-serif'`,
		// which every binding's segment-font resolver passes straight into
		// `get()`/the outline lookup - it must still match the plain family
		// name embedded fonts are registered under.
		const cache = new GlyphOutlineFontCache();
		cache.registerEmbeddedFonts([
			{ name: 'WarpOutlineTestFont', dataUrl: '', rawFontData: testFontBytes() },
		]);
		expect(
			cache.get('"WarpOutlineTestFont", "Some Fallback", sans-serif', false, false),
		).toBeDefined();
		expect(cache.get('WarpOutlineTestFont, sans-serif', false, false)).toBeDefined();
	});

	it('decodes and parses an embedded font only when its variant is looked up', () => {
		const cache = new GlyphOutlineFontCache();
		const reads: string[] = [];
		/** A font whose data reads are counted, like a deferred core descriptor. */
		const counted = (bold: boolean): PptxEmbeddedFont => ({
			name: 'WarpOutlineTestFont',
			bold,
			get dataUrl() {
				reads.push(`dataUrl:${bold}`);
				return '';
			},
			get rawFontData() {
				reads.push(`rawFontData:${bold}`);
				return testFontBytes();
			},
		});
		cache.registerEmbeddedFonts([counted(false), counted(true)]);
		expect(reads).toStrictEqual([]);
		expect(cache.get('WarpOutlineTestFont', false, false)).toBeDefined();
		expect(new Set(reads)).toStrictEqual(new Set(['dataUrl:false', 'rawFontData:false']));
		const afterFirst = reads.length;
		cache.get('WarpOutlineTestFont', false, false);
		expect(reads).toHaveLength(afterFirst);
		expect(cache.get('WarpOutlineTestFont', true, false)).toBeDefined();
		expect(reads.slice(afterFirst).every((read) => read.endsWith(':true'))).toBe(true);
	});

	it('keeps the first registered face for a variant, as eager registration did', () => {
		const cache = new GlyphOutlineFontCache();
		const first: PptxEmbeddedFont = {
			name: 'WarpOutlineTestFont',
			dataUrl: '',
			rawFontData: testFontBytes(),
		};
		const second: PptxEmbeddedFont = {
			name: 'WarpOutlineTestFont',
			dataUrl: '',
			rawFontData: new Uint8Array([1, 2, 3, 4, 5]),
		};
		cache.registerEmbeddedFonts([first]);
		cache.registerEmbeddedFonts([second]);
		expect(cache.get('WarpOutlineTestFont', false, false)).toBeDefined();
	});

	it('keeps a registered embedded face when later webfont bytes do not parse', () => {
		const cache = new GlyphOutlineFontCache();
		cache.registerEmbeddedFonts([
			{ name: 'WarpOutlineTestFont', dataUrl: '', rawFontData: testFontBytes() },
		]);
		expect(cache.registerFontBytes('WarpOutlineTestFont', false, false, new Uint8Array(8))).toBe(
			false,
		);
		expect(cache.get('WarpOutlineTestFont', false, false)).toBeDefined();
	});

	it('registerFontBytes parses and caches directly (the webfont-fetch path)', () => {
		const cache = new GlyphOutlineFontCache();
		const ok = cache.registerFontBytes('WarpOutlineTestFont', false, false, testFontBytes());
		expect(ok).toBeTruthy();
		expect(cache.get('WarpOutlineTestFont', false, false)).toBeDefined();
	});
});

describe('createGlyphOutlineLookup', () => {
	it("composes cache lookup + outline extraction into buildGlyphEnvelope's callback shape", () => {
		const cache = new GlyphOutlineFontCache();
		cache.registerEmbeddedFonts([
			{ name: 'WarpOutlineTestFont', dataUrl: '', rawFontData: testFontBytes() },
		]);
		const lookup = createGlyphOutlineLookup(cache);
		const commands = lookup('A', { fontFamily: 'WarpOutlineTestFont', fontSizePx: 1000 }, 0, 0);
		expect(commands).toBeDefined();
		expect(commands!.length).toBeGreaterThan(0);
	});

	it('returns undefined when no font was registered for the family', () => {
		const cache = new GlyphOutlineFontCache();
		const lookup = createGlyphOutlineLookup(cache);
		expect(lookup('A', { fontFamily: 'Unregistered' }, 0, 0)).toBeUndefined();
	});

	it('answers whitespace with an empty outline even without a font file', () => {
		const lookup = createGlyphOutlineLookup(new GlyphOutlineFontCache());
		expect(lookup(' ', { fontFamily: 'Unregistered' }, 0, 0)).toStrictEqual([]);
	});
});
