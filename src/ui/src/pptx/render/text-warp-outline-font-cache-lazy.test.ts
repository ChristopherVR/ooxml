import { Font, Glyph, Path, parse as parseFont } from 'opentype.js';
import { afterEach, describe, expect, it, vi } from 'vitest';

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
