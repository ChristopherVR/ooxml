import { Font, Glyph, Path } from 'opentype.js';
import type { PptxEmbeddedFont } from 'ooxml-core/pptx';
import { beforeAll, describe, expect, it, vi } from 'vitest';

import {
	createEmbeddedOutlineFontsSlot,
	type EmbeddedFontCandidate,
} from './glyph-outline-font-registration';
import { GlyphOutlineFontCache } from './text-warp-outline-font-cache';
import { preloadOpenTypeFontParser } from './text-warp-opentype-parser';

/** A TrueType file with `.notdef` and one glyph for `char`, at OS/2 weight `weight`. */
function fontBytes(char: string, weight = 400): Uint8Array {
	const notdef = new Glyph({ name: '.notdef', unicode: 0, advanceWidth: 650, path: new Path() });
	const path = new Path();
	path.moveTo(100, 0);
	path.lineTo(100, 700);
	path.lineTo(300, 700);
	path.close();
	const glyph = new Glyph({ name: char, unicode: char.charCodeAt(0), advanceWidth: 650, path });
	const font = new Font({
		familyName: 'RegistrationTest',
		styleName: 'Regular',
		unitsPerEm: 1000,
		ascender: 800,
		descender: -200,
		weightClass: weight,
		glyphs: [notdef, glyph],
	});
	return new Uint8Array(font.toArrayBuffer());
}

/** The candidates `cache` still holds, the way the review measured them. */
function retained(cache: GlyphOutlineFontCache): {
	candidates: number;
	lists: number;
	bytes: number;
} {
	const pending = (cache as unknown as { pendingEmbedded: Iterable<EmbeddedFontCandidate> })
		.pendingEmbedded;
	const all = [...pending];
	return {
		candidates: all.length,
		lists: new Set(all.map((candidate) => candidate.list)).size,
		bytes: all.reduce((sum, candidate) => sum + (sizes.get(candidate.font) ?? 0), 0),
	};
}

// These cases cover registration with the parser loaded;
// text-warp-outline-font-cache-lazy.test.ts covers releasing before it loads.
beforeAll(() => preloadOpenTypeFontParser());

/** Part sizes of the fonts made by {@link deck}, so counting them reads no font data. */
const sizes = new WeakMap<PptxEmbeddedFont, number>();

/**
 * One document's font list: a regular `name` face of `size` bytes whose data
 * reads are counted in `reads`, like a deferred core descriptor.
 */
function deck(name: string, size: number, reads: { count: number }): PptxEmbeddedFont[] {
	const bytes = new Uint8Array(size);
	const font: PptxEmbeddedFont = {
		name,
		get dataUrl() {
			reads.count++;
			return '';
		},
		get rawFontData() {
			reads.count++;
			return bytes;
		},
	};
	sizes.set(font, size);
	return [font];
}

describe('embedded fonts registered with the glyph-outline cache', () => {
	it('are released with their documents without being decoded', () => {
		const cache = new GlyphOutlineFontCache();
		const reads = { count: 0 };
		const releases = Array.from({ length: 100 }, () =>
			cache.registerEmbeddedFonts(deck('Lato', 64 * 1024, reads)),
		);
		expect(retained(cache)).toStrictEqual({ candidates: 100, lists: 100, bytes: 6_553_600 });

		for (const release of releases.slice(0, 50)) {
			release();
		}
		expect(retained(cache)).toStrictEqual({ candidates: 50, lists: 50, bytes: 3_276_800 });
		for (const release of releases.slice(50)) {
			release();
		}
		expect(retained(cache)).toStrictEqual({ candidates: 0, lists: 0, bytes: 0 });
		expect(reads.count).toBe(0);
		expect(cache.hasAttempted('Lato', false, false)).toBe(false);
	});

	it('are recorded once however often the same list is registered', () => {
		const cache = new GlyphOutlineFontCache();
		const list = deck('Lato', 1024, { count: 0 });
		const first = cache.registerEmbeddedFonts(list);
		const second = cache.registerEmbeddedFonts(list);
		const third = cache.registerEmbeddedFonts(list);
		expect(retained(cache)).toStrictEqual({ candidates: 1, lists: 1, bytes: 1024 });

		// Each registration is an owner: the records stay until the last one releases,
		// and releasing the same handle twice counts once.
		second();
		second();
		first();
		expect(retained(cache)).toStrictEqual({ candidates: 1, lists: 1, bytes: 1024 });
		third();
		expect(retained(cache).candidates).toBe(0);

		// A handle from before the list was fully released does not undo a later registration.
		const later = cache.registerEmbeddedFonts(list);
		first();
		third();
		expect(retained(cache)).toStrictEqual({ candidates: 1, lists: 1, bytes: 1024 });
		later();
		expect(retained(cache).candidates).toBe(0);
	});

	it('keep the first trusted candidate first, among the documents still open', () => {
		const face = (bytes: Uint8Array): PptxEmbeddedFont[] => [
			{ name: 'Lato', dataUrl: '', rawFontData: bytes },
		];
		// Lato is served by Google Fonts, so a regular slot holding a weight-800
		// file is not trusted and is passed over.
		const untrusted = face(fontBytes('U', 800));
		const firstTrusted = face(fontBytes('A'));
		const secondTrusted = face(fontBytes('B'));

		const open = new GlyphOutlineFontCache();
		for (const list of [untrusted, firstTrusted, secondTrusted]) {
			open.registerEmbeddedFonts(list);
		}
		expect(open.get('Lato', false, false)?.hasChar('A')).toBe(true);

		const closed = new GlyphOutlineFontCache();
		closed.registerEmbeddedFonts(untrusted);
		const releaseFirst = closed.registerEmbeddedFonts(firstTrusted);
		closed.registerEmbeddedFonts(secondTrusted);
		releaseFirst();
		expect(closed.get('Lato', false, false)?.hasChar('B')).toBe(true);
	});

	it('keep a font parsed before its document was released', () => {
		const cache = new GlyphOutlineFontCache();
		const release = cache.registerEmbeddedFonts([
			{ name: 'RegistrationTest', dataUrl: '', rawFontData: fontBytes('A') },
		]);
		expect(cache.get('RegistrationTest', false, false)).toBeDefined();
		release();
		expect(cache.get('RegistrationTest', false, false)?.hasChar('A')).toBe(true);
	});
});

describe('createEmbeddedOutlineFontsSlot', () => {
	function spyCache() {
		const cache = new GlyphOutlineFontCache();
		const releases: ReturnType<typeof vi.fn>[] = [];
		const register = vi.fn((fonts: readonly PptxEmbeddedFont[]) => {
			const release = vi.fn(cache.registerEmbeddedFonts(fonts));
			releases.push(release);
			return release;
		});
		return { cache, register, releases, registrar: { registerEmbeddedFonts: register } };
	}

	it('releases the previous deck when another is set, and the current one on release', () => {
		const { cache, register, releases, registrar } = spyCache();
		const slot = createEmbeddedOutlineFontsSlot(registrar);
		const first = deck('Lato', 10, { count: 0 });
		const second = deck('Lato', 20, { count: 0 });

		slot.set(first);
		slot.set(first);
		expect(register).toHaveBeenCalledTimes(1);
		expect(retained(cache)).toStrictEqual({ candidates: 1, lists: 1, bytes: 10 });

		slot.set(second);
		expect(releases[0]).toHaveBeenCalledTimes(1);
		expect(retained(cache)).toStrictEqual({ candidates: 1, lists: 1, bytes: 20 });

		slot.release();
		slot.release();
		expect(releases[1]).toHaveBeenCalledTimes(1);
		expect(retained(cache)).toStrictEqual({ candidates: 0, lists: 0, bytes: 0 });
	});

	it('releases on an empty or missing list without registering it', () => {
		const { cache, register, releases, registrar } = spyCache();
		const slot = createEmbeddedOutlineFontsSlot(registrar);
		slot.set(deck('Lato', 10, { count: 0 }));
		slot.set([]);
		expect(releases[0]).toHaveBeenCalledTimes(1);
		slot.set(null);
		slot.set(undefined);
		expect(register).toHaveBeenCalledTimes(1);
		expect(retained(cache).candidates).toBe(0);
	});

	it('keeps viewers sharing one cache independent', () => {
		const cache = new GlyphOutlineFontCache();
		const one = createEmbeddedOutlineFontsSlot(cache);
		const two = createEmbeddedOutlineFontsSlot(cache);
		one.set(deck('Lato', 10, { count: 0 }));
		two.set(deck('Lato', 20, { count: 0 }));
		one.release();
		expect(retained(cache)).toStrictEqual({ candidates: 1, lists: 1, bytes: 20 });
	});

	it('keeps a font list shared by two viewers until both release it', () => {
		const cache = new GlyphOutlineFontCache();
		const shared: PptxEmbeddedFont[] = [
			{ name: 'RegistrationTest', dataUrl: '', rawFontData: fontBytes('A') },
		];
		const one = createEmbeddedOutlineFontsSlot(cache);
		const two = createEmbeddedOutlineFontsSlot(cache);
		one.set(shared);
		two.set(shared);
		expect(retained(cache).candidates).toBe(1);

		// The first viewer closes; the second still shows the same list and needs its outline.
		one.release();
		expect(retained(cache).candidates).toBe(1);
		two.set(shared);
		expect(cache.get('RegistrationTest', false, false)?.hasChar('A')).toBe(true);
		two.release();
		expect(retained(cache).candidates).toBe(0);
	});

	it('releases a shared list once its last viewer moves to another deck', () => {
		const cache = new GlyphOutlineFontCache();
		const shared = deck('Lato', 10, { count: 0 });
		const one = createEmbeddedOutlineFontsSlot(cache);
		const two = createEmbeddedOutlineFontsSlot(cache);
		one.set(shared);
		two.set(shared);
		one.set(deck('Lato', 20, { count: 0 }));
		expect(retained(cache)).toStrictEqual({ candidates: 2, lists: 2, bytes: 30 });
		two.release();
		expect(retained(cache)).toStrictEqual({ candidates: 1, lists: 1, bytes: 20 });
		one.release();
		expect(retained(cache).candidates).toBe(0);
	});
});
