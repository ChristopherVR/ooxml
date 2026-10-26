import { Font, Glyph, Path, parse as parseFont } from 'opentype.js';
import type { PptxEmbeddedFont } from 'ooxml-core/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { GlyphOutlineFontCache } from './text-warp-outline-font-cache';

const path = new Path();
path.moveTo(100, 0);
path.lineTo(100, 700);
path.lineTo(300, 700);
path.close();
const font = new Font({
	familyName: 'LazyOutline',
	styleName: 'Regular',
	unitsPerEm: 1000,
	ascender: 800,
	descender: -200,
	glyphs: [
		new Glyph({ name: '.notdef', unicode: 0, advanceWidth: 650, path: new Path() }),
		new Glyph({ name: 'A', unicode: 65, advanceWidth: 650, path }),
	],
});
const bytes = new Uint8Array(font.toArrayBuffer());
const embedded = [{ name: 'LazyOutline', dataUrl: '', rawFontData: bytes }];
const spec = { fontFamily: 'LazyOutline', fontSizePx: 1000 };

afterEach(() => {
	vi.doUnmock('opentype.js');
	vi.resetModules();
});

/** Fresh cache and parser modules over a mocked `opentype.js` whose import waits for `release`. */
async function freshCache() {
	vi.resetModules();
	let release!: () => void;
	const gate = new Promise<void>((resolve) => {
		release = resolve;
	});
	const parse = vi.fn(parseFont);
	const factory = vi.fn(async () => {
		// Bounded, so a regression to a static import fails instead of hanging the import.
		await Promise.race([
			gate,
			new Promise<void>((resolve) => {
				setTimeout(resolve, 1000);
			}),
		]);
		return { parse: undefined, default: { parse } };
	});
	vi.doMock('opentype.js', factory);
	const { GlyphOutlineFontCache, createGlyphOutlineLookup } =
		await import('./text-warp-outline-font-cache');
	const parser = await import('./text-warp-opentype-parser');
	const cache = new GlyphOutlineFontCache();
	return { cache, lookup: createGlyphOutlineLookup(cache), parser, factory, parse, release };
}

/** A one-glyph font drawing `char`, so a test can tell which candidate a lookup resolved to. */
function fontBytesWithGlyph(char: string): Uint8Array {
	const glyphPath = new Path();
	glyphPath.moveTo(0, 0);
	glyphPath.lineTo(0, 500);
	glyphPath.lineTo(500, 500);
	glyphPath.close();
	return new Uint8Array(
		new Font({
			familyName: `Outline${char}`,
			styleName: 'Regular',
			unitsPerEm: 1000,
			ascender: 800,
			descender: -200,
			glyphs: [
				new Glyph({ name: '.notdef', unicode: 0, advanceWidth: 500, path: new Path() }),
				new Glyph({
					name: char,
					unicode: char.codePointAt(0)!,
					advanceWidth: 500,
					path: glyphPath,
				}),
			],
		}).toArrayBuffer(),
	);
}

const fontA = fontBytesWithGlyph('A');
const fontB = fontBytesWithGlyph('B');
/** A WOFF2 header: a format the bundled opentype.js rejects, as Google Fonts serves it. */
const woff2 = new Uint8Array(64);
woff2.set([0x77, 0x4f, 0x46, 0x32, 0x00, 0x01, 0x00, 0x00]);
/** Bytes that are no font at all. */
const garbage = new Uint8Array(64).fill(0x5a);

type Step = { embedded: string; bytes: Uint8Array } | { webfont: string; bytes: Uint8Array };

function apply(cache: GlyphOutlineFontCache, steps: readonly Step[]): void {
	for (const step of steps) {
		if ('embedded' in step) {
			cache.registerEmbeddedFonts([{ name: step.embedded, dataUrl: '', rawFontData: step.bytes }]);
		} else {
			cache.registerFontBytes(step.webfont, false, false, step.bytes);
		}
	}
}

/** Which test glyph the cache resolves `family` to, if any. */
function resolvedGlyph(cache: GlyphOutlineFontCache, family: string): string | undefined {
	const resolved = cache.get(family, false, false);
	if (!resolved) {
		return undefined;
	}
	return ['A', 'B'].find((char) => resolved.hasChar(char)) ?? 'none';
}

async function settle(): Promise<void> {
	for (let i = 0; i < 5; i += 1) {
		await Promise.resolve();
	}
}

describe('GlyphOutlineFontCache loads opentype.js on demand', () => {
	it('does not load the parser to import the cache or register fonts', async () => {
		const { cache, parser, factory } = await freshCache();
		cache.registerEmbeddedFonts(embedded);
		expect(cache.registerFontBytes('Webfont', false, false, bytes)).toBe(false);
		await settle();
		expect(factory).not.toHaveBeenCalled();
		expect(parser.isOpenTypeFontParserReady()).toBe(false);
		// Waiting bytes are not failed attempts: they still parse once the parser loads.
		expect(cache.hasAttempted('LazyOutline', false, false)).toBe(false);
		expect(cache.hasAttempted('Webfont', false, false)).toBe(false);
	});

	it('does not load the parser for whitespace or for families it holds no bytes for', async () => {
		const { cache, lookup, factory } = await freshCache();
		cache.registerEmbeddedFonts(embedded);
		expect(lookup(' ', spec, 0, 0)).toStrictEqual([]);
		expect(lookup('A', { fontFamily: 'Unregistered' }, 0, 0)).toBeUndefined();
		await settle();
		expect(factory).not.toHaveBeenCalled();
	});

	it('loads the parser once when a lookup needs a waiting font, then notifies and serves its outline', async () => {
		const { cache, lookup, parser, factory, parse, release } = await freshCache();
		const listener = vi.fn();
		const unsubscribe = cache.subscribe(listener);
		cache.registerEmbeddedFonts(embedded);
		// No DOM canvas here, so the browser-trace fallback has nothing to offer.
		expect(lookup('A', spec, 0, 0)).toBeUndefined();
		expect(lookup('A', spec, 0, 0)).toBeUndefined();
		await vi.waitFor(() => expect(factory).toHaveBeenCalledOnce());
		expect(listener).not.toHaveBeenCalled();
		release();
		await cache.preload();
		await settle();
		expect(parser.isOpenTypeFontParserReady()).toBe(true);
		expect(factory).toHaveBeenCalledOnce();
		expect(parse).toHaveBeenCalledOnce();
		expect(listener).toHaveBeenCalledOnce();
		expect(cache.getSnapshot()).toBe(1);
		expect(lookup('A', spec, 0, 0)).toStrictEqual(font.getPath('A', 0, 0, 1000).commands);
		unsubscribe();
		cache.registerFontBytes('Another', false, false, bytes);
		await settle();
		expect(listener).toHaveBeenCalledOnce();
	});

	it('starts the load when bytes arrive for a family a lookup already missed', async () => {
		const { cache, lookup, factory, release } = await freshCache();
		const listener = vi.fn();
		cache.subscribe(listener);
		expect(lookup('A', spec, 0, 0)).toBeUndefined();
		expect(factory).not.toHaveBeenCalled();
		expect(cache.registerFontBytes('LazyOutline', false, false, bytes)).toBe(false);
		await vi.waitFor(() => expect(factory).toHaveBeenCalledOnce());
		release();
		await cache.preload();
		await settle();
		expect(listener).toHaveBeenCalledOnce();
		expect(cache.get('LazyOutline', false, false)?.hasChar('A')).toBe(true);
	});

	it('keeps waiting fonts after a failed import and retries on the next demand', async () => {
		const { cache, release } = await freshCache();
		vi.doMock('opentype.js', () => {
			throw new Error('chunk unavailable');
		});
		cache.registerEmbeddedFonts(embedded);
		await expect(cache.preload()).rejects.toThrow();
		expect(cache.hasAttempted('LazyOutline', false, false)).toBe(false);
		vi.doUnmock('opentype.js');
		release();
		await cache.preload();
		expect(cache.get('LazyOutline', false, false)?.hasChar('A')).toBe(true);
	});

	it('parses synchronously again once the parser is loaded', async () => {
		const { cache, release } = await freshCache();
		release();
		await cache.preload();
		expect(cache.registerFontBytes('LazyOutline', false, false, bytes)).toBe(true);
		cache.registerEmbeddedFonts([{ ...embedded[0]!, name: 'LateDeck' }]);
		expect(cache.get('LateDeck', false, false)?.hasChar('A')).toBe(true);
	});

	it('notifies when a font registered after a missed lookup parses synchronously', async () => {
		const { cache, lookup, release } = await freshCache();
		release();
		await cache.preload();
		const listener = vi.fn();
		cache.subscribe(listener);
		expect(lookup('A', spec, 0, 0)).toBeUndefined();
		cache.registerEmbeddedFonts(embedded);
		cache.registerFontBytes('LazyOutline', true, false, bytes);
		expect(listener).not.toHaveBeenCalled();
		await settle();
		expect(listener).toHaveBeenCalledOnce();
	});
});

describe('GlyphOutlineFontCache keeps the eager precedence between candidates', () => {
	const brand = 'BrandOutline';
	// Expected outcomes are those of parsing every registration synchronously:
	// a later candidate replaces the font only once it parses, an embedded font
	// never replaces one already registered for its key, and keys ignore casing.
	const cases: ReadonlyArray<{ name: string; steps: Step[]; expected: string | undefined }> = [
		{
			name: 'an embedded font survives an unsupported WOFF2 replacement',
			steps: [
				{ embedded: brand, bytes: fontA },
				{ webfont: brand, bytes: woff2 },
			],
			expected: 'A',
		},
		{
			name: 'an embedded font survives an invalid replacement',
			steps: [
				{ embedded: brand, bytes: fontA },
				{ webfont: brand, bytes: garbage },
			],
			expected: 'A',
		},
		{
			name: 'a later webfont that parses replaces the embedded font',
			steps: [
				{ embedded: brand, bytes: fontA },
				{ webfont: brand, bytes: fontB },
			],
			expected: 'B',
		},
		{
			name: 'an embedded font does not replace an earlier webfont',
			steps: [
				{ webfont: brand, bytes: fontB },
				{ embedded: brand, bytes: fontA },
			],
			expected: 'B',
		},
		{
			name: 'an embedded font is not registered after a failed webfont for its key',
			steps: [
				{ webfont: brand, bytes: woff2 },
				{ embedded: brand, bytes: fontA },
			],
			expected: undefined,
		},
		{
			name: 'a webfont survives a later webfont that fails',
			steps: [
				{ webfont: brand, bytes: fontA },
				{ webfont: brand, bytes: woff2 },
			],
			expected: 'A',
		},
		{
			name: 'the last of several webfonts that parse wins',
			steps: [
				{ webfont: brand, bytes: fontA },
				{ webfont: brand, bytes: fontB },
				{ webfont: brand, bytes: woff2 },
			],
			expected: 'B',
		},
		{
			name: 'a failed candidate does not block a later webfont that parses',
			steps: [
				{ embedded: brand, bytes: woff2 },
				{ webfont: brand, bytes: garbage },
				{ webfont: brand, bytes: fontA },
			],
			expected: 'A',
		},
		{
			name: 'the first embedded font wins across family-name casing',
			steps: [
				{ embedded: brand, bytes: fontA },
				{ embedded: brand.toUpperCase(), bytes: fontB },
			],
			expected: 'A',
		},
		{
			name: 'an embedded font survives failed replacements under other casings',
			steps: [
				{ embedded: brand, bytes: fontA },
				{ webfont: brand.toLowerCase(), bytes: woff2 },
				{ webfont: `"${brand.toUpperCase()}", sans-serif`, bytes: garbage },
			],
			expected: 'A',
		},
	];

	for (const { name, steps, expected } of cases) {
		it(`${name} (parser already loaded)`, async () => {
			const { cache, release } = await freshCache();
			release();
			await cache.preload();
			apply(cache, steps);
			expect(resolvedGlyph(cache, brand)).toBe(expected);
		});

		it(`${name} (parser loaded afterwards)`, async () => {
			const { cache, parser, release } = await freshCache();
			apply(cache, steps);
			expect(parser.isOpenTypeFontParserReady()).toBe(false);
			release();
			await cache.preload();
			expect(resolvedGlyph(cache, brand)).toBe(expected);
		});
	}

	it('keeps the embedded font when a lookup loads the parser after a failed WOFF2 replacement', async () => {
		const { cache, lookup, release } = await freshCache();
		const listener = vi.fn();
		cache.subscribe(listener);
		cache.registerEmbeddedFonts([{ name: brand, dataUrl: '', rawFontData: fontA }]);
		expect(cache.registerFontBytes(brand, false, false, woff2)).toBe(false);
		expect(lookup('A', { fontFamily: brand, fontSizePx: 1000 }, 0, 0)).toBeUndefined();
		release();
		await cache.preload();
		await settle();
		expect(listener).toHaveBeenCalledOnce();
		expect(resolvedGlyph(cache, brand)).toBe('A');
	});

	it('parses a file registered again on each deck load only once', async () => {
		const { cache, parse, release } = await freshCache();
		cache.registerEmbeddedFonts([{ name: brand, dataUrl: '', rawFontData: fontA }]);
		for (let load = 0; load < 3; load += 1) {
			cache.registerFontBytes(brand, false, false, woff2.slice());
		}
		release();
		await cache.preload();
		expect(parse).toHaveBeenCalledTimes(2);
		expect(resolvedGlyph(cache, brand)).toBe('A');
	});

	it('keeps a waiting font when a replacement fails after another deck loaded the parser', async () => {
		const { cache, parser, release } = await freshCache();
		const { GlyphOutlineFontCache: Cache } = await import('./text-warp-outline-font-cache');
		cache.registerEmbeddedFonts([{ name: brand, dataUrl: '', rawFontData: fontA }]);
		release();
		await new Cache().preload();
		expect(parser.isOpenTypeFontParserReady()).toBe(true);
		expect(cache.registerFontBytes(brand.toLowerCase(), false, false, woff2)).toBe(false);
		expect(resolvedGlyph(cache, brand)).toBe('A');
		expect(cache.registerFontBytes(brand, false, false, fontB)).toBe(true);
		expect(resolvedGlyph(cache, brand)).toBe('B');
	});
});

describe('GlyphOutlineFontCache decodes deferred embedded fonts only when looked up', () => {
	const brand = 'BrandOutline';

	/** An embedded variant whose data reads are counted, like a deferred core descriptor. */
	function counted(reads: string[], bytes: Uint8Array, bold = false): PptxEmbeddedFont {
		const tag = `${brand}:${bold ? 'bold' : 'regular'}`;
		return {
			name: brand,
			bold,
			get dataUrl() {
				reads.push(tag);
				return '';
			},
			get rawFontData() {
				reads.push(tag);
				return bytes;
			},
		};
	}

	it('decodes only the variant a lookup asked for when that lookup loads the parser', async () => {
		const { cache, lookup, parse, release } = await freshCache();
		const reads: string[] = [];
		const listener = vi.fn();
		cache.subscribe(listener);
		cache.registerEmbeddedFonts([counted(reads, fontA), counted(reads, fontB, true)]);
		expect(lookup('A', { fontFamily: brand, fontSizePx: 1000 }, 0, 0)).toBeUndefined();
		expect(reads).toStrictEqual([]);
		release();
		await vi.waitFor(() => expect(listener).toHaveBeenCalledOnce());
		expect(new Set(reads)).toStrictEqual(new Set([`${brand}:regular`]));
		expect(parse).toHaveBeenCalledOnce();
		expect(resolvedGlyph(cache, brand)).toBe('A');
		expect(cache.get(brand, true, false)?.hasChar('B')).toBe(true);
		expect(reads).toContain(`${brand}:bold`);
	});

	it('never decodes an embedded font that a waiting webfont replaces', async () => {
		const { cache, lookup, release } = await freshCache();
		const reads: string[] = [];
		const listener = vi.fn();
		cache.subscribe(listener);
		cache.registerEmbeddedFonts([counted(reads, fontA)]);
		expect(cache.registerFontBytes(brand, false, false, fontB)).toBe(false);
		expect(lookup('B', { fontFamily: brand, fontSizePx: 1000 }, 0, 0)).toBeUndefined();
		release();
		await vi.waitFor(() => expect(listener).toHaveBeenCalledOnce());
		expect(resolvedGlyph(cache, brand)).toBe('B');
		expect(reads).toStrictEqual([]);
	});

	it('decodes the embedded font once a failing webfont leaves it as the answer', async () => {
		const { cache, lookup, release } = await freshCache();
		const reads: string[] = [];
		const listener = vi.fn();
		cache.subscribe(listener);
		cache.registerEmbeddedFonts([counted(reads, fontA)]);
		expect(cache.registerFontBytes(brand, false, false, woff2)).toBe(false);
		expect(lookup('A', { fontFamily: brand, fontSizePx: 1000 }, 0, 0)).toBeUndefined();
		expect(reads).toStrictEqual([]);
		release();
		await vi.waitFor(() => expect(listener).toHaveBeenCalledOnce());
		expect(resolvedGlyph(cache, brand)).toBe('A');
		expect(reads.length).toBeGreaterThan(0);
	});
});
