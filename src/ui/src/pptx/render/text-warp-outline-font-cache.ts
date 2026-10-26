/**
 * Best-effort glyph-outline font source for `text-warp-glyph-outline.ts`.
 *
 * Parses actual font FILES (TrueType/CFF) with `opentype.js` so a WordArt
 * envelope glyph can be warped as its real outline instead of an affine
 * approximation of its bounding box (see that module's doc comment). Outlines
 * are obtainable from two places, matching CLAUDE.md's Rule 2 "pure decision
 * function" shape: this module only ever produces a `Font | undefined`
 * lookup, never touches the DOM, and never decides how a binding renders.
 *
 *  - Embedded fonts (`PptxEmbeddedFont[]`): the deck's own loader already
 *    holds the clear-text bytes in memory (see
 *    `embedded-fonts.ts`'s `resolveEmbeddedFontClearBytes`), so parsing is
 *    synchronous and needs no network round-trip.
 *  - Catalogue webfonts (Google Fonts): the family's actual `.woff2` bytes
 *    are fetched separately from the `<link rel=stylesheet>` each binding
 *    already injects for on-screen text (see
 *    `text-warp-outline-webfont-fetch.ts`), since the stylesheet alone
 *    carries no usable glyph geometry.
 *
 * A font this cache cannot resolve (parse failure, unknown family, a webfont
 * fetch still in flight, or a system font with no file the browser exposes at
 * all) simply has no entry: {@link createGlyphOutlineLookup} then traces the
 * glyph the browser renders instead (`text-warp-glyph-trace.ts`), and only
 * without a DOM canvas does the layout fall back to an affine fit.
 *
 * `opentype.js` itself is loaded on demand (`text-warp-opentype-parser.ts`):
 * registered bytes wait in the cache until an envelope glyph asks for their
 * family, which starts the one shared import. Until it lands that glyph is
 * traced as above; the cache then parses the waiting fonts and notifies its
 * {@link GlyphOutlineFontCache.subscribe} listeners so bindings repaint with
 * the file's outlines. Decks without envelope WordArt never load the parser.
 */
import type { Font } from 'opentype.js';
import type { PptxEmbeddedFont } from 'ooxml-core/pptx';

import { resolveEmbeddedFontClearBytes } from './embedded-font-bytes';
import { selectTrustedEmbeddedFonts } from './embedded-font-trust';
import {
	isOpenTypeFontParserReady,
	parseOpenTypeFont,
	preloadOpenTypeFontParser,
} from './text-warp-opentype-parser';

export {
	createGlyphOutlineLookup,
	extractGlyphOutlineCommands,
	type GlyphOutlineCacheFontSpec,
} from './text-warp-glyph-outline-lookup';

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
	// `opentype.parse` wants a real ArrayBuffer; `bytes` may be a view over a
	// larger buffer (e.g. `rawFontData` sliced from a ZIP entry), so this
	// copies only the bytes this font actually owns.
	const copy = bytes.slice();
	return copy.buffer as ArrayBuffer;
}

/**
 * The primary family name out of a CSS `font-family` VALUE, which may be a
 * comma-separated fallback list with quoted entries (e.g.
 * `'"Calibri", "Carlito", sans-serif'`, what `getSubstituteFontFamily` in
 * `ooxml-core/pptx` returns and every binding's segment-font resolver
 * passes through here) - a plain, single family name (what
 * `PptxEmbeddedFont.name` and `collectReferencedFontFamilies` already
 * produce) passes through unchanged.
 */
function primaryFontFamily(family: string): string {
	const first = family.split(',')[0] ?? family;
	return first.trim().replace(/^['"]|['"]$/gu, '');
}

function cacheKey(family: string, bold: boolean, italic: boolean): string {
	return `${primaryFontFamily(family).toLowerCase()}|${bold ? 1 : 0}|${italic ? 1 : 0}`;
}

/**
 * Caches parsed `opentype.js` fonts by `(family, bold, italic)`.
 *
 * A binding constructs one instance per loaded deck, calls
 * {@link registerEmbeddedFonts} as soon as the deck's embedded fonts are
 * known, optionally feeds it webfont bytes as they arrive via
 * {@link registerFontBytes}, passes `.get(...)` into `buildGlyphEnvelope`'s
 * `getOutlineFont` option, and repaints when a {@link subscribe} listener
 * fires. Registration parses synchronously once the parser is loaded; before
 * that, valid bytes are kept and parsed when an outline lookup needs them.
 * Listeners fire (once per microtask) when a font is parsed for a family a
 * lookup already missed, so a glyph drawn by the fallback gets its outline.
 */
export class GlyphOutlineFontCache {
	private readonly fonts = new Map<string, Font>();
	private readonly attempted = new Set<string>();
	/** Registered bytes waiting for the parser to load. */
	private readonly pending = new Map<string, Uint8Array>();
	/** Keys a lookup missed: parsing one later changes what a repaint draws. */
	private readonly missed = new Set<string>();
	private readonly listeners = new Set<() => void>();
	private revision = 0;
	private notifyQueued = false;
	private loading: Promise<void> | undefined;

	/**
	 * Call `listener` when a newly parsed font makes an outline available that
	 * an earlier lookup could not use. Returns the unsubscribe function. Shaped
	 * for `useSyncExternalStore`.
	 */
	readonly subscribe = (listener: () => void): (() => void) => {
		this.listeners.add(listener);
		return () => {
			this.listeners.delete(listener);
		};
	};

	/** A number that changes each time {@link subscribe} listeners are notified. */
	readonly getSnapshot = (): number => this.revision;

	/**
	 * Load the parser if needed and parse every waiting font, so later lookups
	 * resolve synchronously. Await it before a render that must use font-file
	 * outlines on its first pass. Rejects if the parser cannot be loaded; the
	 * waiting fonts are kept, so a later call retries.
	 */
	preload(): Promise<void> {
		if (isOpenTypeFontParserReady()) {
			this.parsePending();
			return Promise.resolve();
		}
		this.loading ??= preloadOpenTypeFontParser().then(
			() => {
				this.loading = undefined;
				this.parsePending();
			},
			(error: unknown) => {
				this.loading = undefined;
				throw error;
			},
		);
		return this.loading;
	}

	/**
	 * The parsed font for `(family, bold, italic)`, or `undefined` if unavailable.
	 * Missing a font whose bytes wait for the parser starts loading it.
	 */
	get(
		family: string | undefined,
		bold: boolean | undefined,
		italic: boolean | undefined,
	): Font | undefined {
		if (!family) {
			return undefined;
		}
		const key = cacheKey(family, Boolean(bold), Boolean(italic));
		const font = this.fonts.get(key);
		if (font) {
			return font;
		}
		const bytes = this.pending.get(key);
		if (bytes && isOpenTypeFontParserReady()) {
			this.pending.delete(key);
			this.parseInto(key, bytes);
			return this.fonts.get(key);
		}
		this.missed.add(key);
		if (bytes) {
			this.requestParser();
		}
		return undefined;
	}

	/** True once a parse of this `(family, bold, italic)` was attempted, whether or not it resolved. */
	hasAttempted(family: string, bold: boolean, italic: boolean): boolean {
		return this.attempted.has(cacheKey(family, bold, italic));
	}

	/**
	 * Parse every embedded font whose bytes are already resolvable, no network
	 * involved: synchronously once the parser is loaded, otherwise when an
	 * outline lookup first needs one of them. Safe to call repeatedly (e.g.
	 * once per render) with the same list: already-resolved, already-failed and
	 * waiting entries are skipped. Typefaces whose files do not match them are
	 * skipped too (see `selectTrustedEmbeddedFonts`), exactly as for `@font-face`.
	 */
	registerEmbeddedFonts(fonts: readonly PptxEmbeddedFont[]): void {
		const ready = isOpenTypeFontParserReady();
		let needed = false;
		for (const font of selectTrustedEmbeddedFonts(fonts)) {
			if (!font.name) {
				continue;
			}
			const key = cacheKey(font.name, Boolean(font.bold), Boolean(font.italic));
			if (this.fonts.has(key) || this.attempted.has(key) || this.pending.has(key)) {
				continue;
			}
			const bytes = resolveEmbeddedFontClearBytes(font);
			if (!bytes || bytes.length < 4) {
				this.attempted.add(key);
				continue;
			}
			if (ready) {
				// Unparseable fonts (corrupt, or an outline format opentype.js does
				// not support) stay unset so the traced fallback applies.
				this.parseInto(key, bytes);
			} else {
				this.pending.set(key, bytes);
				needed ||= this.missed.has(key);
			}
		}
		if (needed) {
			this.requestParser();
		}
	}

	/**
	 * Register raw bytes fetched separately for `(family, bold, italic)` (see
	 * `text-warp-outline-webfont-fetch.ts`). Returns whether parsing
	 * succeeded; a caller that owns re-rendering can use this to know whether
	 * a re-render is worth triggering. Before the parser is loaded the bytes
	 * are kept and `false` is returned; {@link subscribe} listeners are told
	 * if they later fill in an outline that a lookup already asked for.
	 */
	registerFontBytes(family: string, bold: boolean, italic: boolean, bytes: Uint8Array): boolean {
		const key = cacheKey(family, bold, italic);
		if (bytes.length < 4) {
			this.attempted.add(key);
			return false;
		}
		if (!isOpenTypeFontParserReady()) {
			this.pending.set(key, bytes);
			if (this.missed.has(key)) {
				this.requestParser();
			}
			return false;
		}
		this.pending.delete(key);
		return this.parseInto(key, bytes);
	}

	private parseInto(key: string, bytes: Uint8Array): boolean {
		this.attempted.add(key);
		try {
			this.fonts.set(key, parseOpenTypeFont(toArrayBuffer(bytes)));
		} catch {
			return false;
		}
		if (this.missed.delete(key)) {
			this.notify();
		}
		return true;
	}

	private parsePending(): void {
		const waiting = [...this.pending];
		this.pending.clear();
		for (const [key, bytes] of waiting) {
			this.parseInto(key, bytes);
		}
	}

	/** Coalesce notifications, and never call listeners from inside a caller's own update. */
	private notify(): void {
		if (this.notifyQueued) {
			return;
		}
		this.notifyQueued = true;
		queueMicrotask(() => {
			this.notifyQueued = false;
			this.revision += 1;
			for (const listener of [...this.listeners]) {
				listener();
			}
		});
	}

	private requestParser(): void {
		// A failed import keeps the traced fallback; the next demand retries it.
		this.preload().catch(() => undefined);
	}
}
