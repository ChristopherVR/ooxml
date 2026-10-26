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
 *  - Embedded fonts (`PptxEmbeddedFont[]`): the deck's own loader holds
 *    the font parts in memory and decodes a variant on first read (see
 *    `embedded-fonts.ts`'s `resolveEmbeddedFontClearBytes`), so parsing is
 *    synchronous and needs no network round-trip. It happens on the first
 *    lookup of that variant.
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
 * registered fonts wait in the cache until an envelope glyph asks for their
 * family, which starts the one shared import. Until it lands that glyph is
 * traced as above; the cache then parses the fonts lookups asked for and
 * notifies its {@link GlyphOutlineFontCache.subscribe} listeners so bindings
 * repaint with the file's outlines. Decks without envelope WordArt never load
 * the parser, and never decode an embedded font for it.
 */
import type { Font } from 'opentype.js';
import type { PptxEmbeddedFont } from 'ooxml-core/pptx';

import { resolveEmbeddedFontClearBytes } from './embedded-font-bytes';
import { isEmbeddedTypefaceTrusted } from './embedded-font-trust';
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

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
	if (a === b) {
		return true;
	}
	if (a.length !== b.length) {
		return false;
	}
	for (let i = 0; i < a.length; i += 1) {
		if (a[i] !== b[i]) {
			return false;
		}
	}
	return true;
}

function cacheKey(family: string, bold: boolean, italic: boolean): string {
	return `${primaryFontFamily(family).toLowerCase()}|${bold ? 1 : 0}|${italic ? 1 : 0}`;
}

/** An embedded variant registered for a key, with the list it came from. */
interface PendingEmbeddedFont {
	font: PptxEmbeddedFont;
	list: readonly PptxEmbeddedFont[];
}

/**
 * Caches parsed `opentype.js` fonts by `(family, bold, italic)`.
 *
 * A binding constructs one instance per loaded deck, calls
 * {@link registerEmbeddedFonts} as soon as the deck's embedded fonts are
 * known, optionally feeds it webfont bytes as they arrive via
 * {@link registerFontBytes}, passes `.get(...)` into `buildGlyphEnvelope`'s
 * `getOutlineFont` option, and repaints when a {@link subscribe} listener
 * fires.
 *
 * Registering an embedded font only records it: its bytes are decoded, its
 * typeface judged (see `isEmbeddedTypefaceTrusted`) and its file parsed the
 * first time its `(family, bold, italic)` is looked up, which only text that
 * draws glyph outlines does. Webfont bytes parse on arrival once the parser
 * is loaded; before that they are kept with the key's other candidates.
 * Either way, lookups answer exactly as if every registered font had been
 * parsed on arrival. Listeners fire (once per microtask) when a font is
 * parsed for a family a lookup already missed, so a glyph drawn by the
 * fallback gets its outline.
 */
export class GlyphOutlineFontCache {
	private readonly fonts = new Map<string, Font>();
	private readonly attempted = new Set<string>();
	/**
	 * Registered embedded candidates per key, in registration order, with
	 * nothing read from them yet. A key's embedded candidates are always older
	 * than its {@link pendingBytes}: an embedded font is not recorded for a key
	 * that already has bytes, as it would not have been parsed on arrival.
	 */
	private readonly pendingEmbedded = new Map<string, PendingEmbeddedFont[]>();
	/**
	 * Webfont bytes registered before the parser loaded, oldest first. Every
	 * candidate is kept until one parses: a later registration only replaces an
	 * earlier one if it parses, exactly as when each is parsed on arrival.
	 */
	private readonly pendingBytes = new Map<string, Uint8Array[]>();
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
	 * Load the parser if needed and parse every waiting font, embedded ones
	 * included, so later lookups resolve synchronously. Await it before a
	 * render that must use font-file outlines on its first pass. Rejects if the
	 * parser cannot be loaded; the waiting fonts are kept, so a later call
	 * retries. The load a lookup starts on its own parses only the fonts that
	 * lookups asked for, so other embedded fonts stay undecoded.
	 */
	preload(): Promise<void> {
		if (isOpenTypeFontParserReady()) {
			this.resolveAllPending();
			return Promise.resolve();
		}
		return this.loadParser().then(() => {
			this.resolveAllPending();
		});
	}

	/**
	 * The parsed font for `(family, bold, italic)`, or `undefined` if unavailable.
	 * Missing a font whose candidates wait for the parser starts loading it.
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
		const waiting = this.hasPending(key);
		if (waiting && isOpenTypeFontParserReady()) {
			this.resolvePending(key);
			return this.fonts.get(key);
		}
		this.missed.add(key);
		if (waiting) {
			this.requestParser();
		}
		return undefined;
	}

	/**
	 * True once a parse of this `(family, bold, italic)` was attempted, whether
	 * or not it resolved. With the parser loaded this settles the key's waiting
	 * candidates first, like a lookup; before that they are not attempts yet.
	 */
	hasAttempted(family: string, bold: boolean, italic: boolean): boolean {
		const key = cacheKey(family, bold, italic);
		if (isOpenTypeFontParserReady()) {
			this.resolvePending(key);
		}
		return this.attempted.has(key);
	}

	/**
	 * Record every embedded font for outline lookups, no network involved and
	 * nothing decoded yet. Safe to call repeatedly (e.g. once per render) with
	 * the same list: already-resolved and already-failed entries are skipped.
	 * Typefaces whose files do not match them are skipped too (see
	 * `isEmbeddedTypefaceTrusted`), exactly as for `@font-face`. A font for a
	 * key a lookup already missed is parsed now if the parser is loaded, and
	 * otherwise starts loading it.
	 */
	registerEmbeddedFonts(fonts: readonly PptxEmbeddedFont[]): void {
		let needed = false;
		for (const font of fonts) {
			if (!font.name) {
				continue;
			}
			const key = cacheKey(font.name, Boolean(font.bold), Boolean(font.italic));
			if (this.fonts.has(key) || this.attempted.has(key) || this.pendingBytes.has(key)) {
				continue;
			}
			const candidates = this.pendingEmbedded.get(key);
			if (candidates) {
				candidates.push({ font, list: fonts });
			} else {
				this.pendingEmbedded.set(key, [{ font, list: fonts }]);
			}
			needed ||= this.missed.has(key);
		}
		if (needed) {
			this.resolveMissedOrLoad();
		}
	}

	/**
	 * Register raw bytes fetched separately for `(family, bold, italic)` (see
	 * `text-warp-outline-webfont-fetch.ts`). Returns whether parsing
	 * succeeded; a caller that owns re-rendering can use this to know whether
	 * a re-render is worth triggering. Before the parser is loaded the bytes
	 * are kept and `false` is returned; {@link subscribe} listeners are told
	 * if they later fill in an outline that a lookup already asked for. Either
	 * way, bytes that do not parse leave the font already registered for this
	 * key (or still waiting for the parser) in place.
	 */
	registerFontBytes(family: string, bold: boolean, italic: boolean, bytes: Uint8Array): boolean {
		const key = cacheKey(family, bold, italic);
		if (bytes.length < 4) {
			this.attempted.add(key);
			return false;
		}
		if (!isOpenTypeFontParserReady()) {
			this.addPendingBytes(key, bytes);
			if (this.missed.has(key)) {
				this.requestParser();
			}
			return false;
		}
		// These bytes are the newest candidate for the key: if they parse they win
		// over anything still waiting (another cache may have loaded the parser
		// while this one held candidates), which is then dropped undecoded. If
		// not, the waiting candidates settle as if each had been parsed on arrival.
		if (this.parseInto(key, bytes)) {
			this.dropPending(key);
			return true;
		}
		this.resolvePending(key);
		return false;
	}

	/** Parse `bytes` for `key`; on failure the font already set for `key` stays. */
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

	private addPendingBytes(key: string, bytes: Uint8Array): void {
		const waiting = this.pendingBytes.get(key) ?? [];
		// The same file registered again (e.g. on every deck load) parses the same
		// way, so only its latest position counts.
		const index = waiting.findIndex((candidate) => sameBytes(candidate, bytes));
		if (index >= 0) {
			waiting.splice(index, 1);
		}
		waiting.push(bytes);
		this.pendingBytes.set(key, waiting);
	}

	private hasPending(key: string): boolean {
		return this.pendingBytes.has(key) || this.pendingEmbedded.has(key);
	}

	private dropPending(key: string): void {
		this.pendingBytes.delete(key);
		this.pendingEmbedded.delete(key);
	}

	/**
	 * Resolve the candidates waiting for `key` as if each had been parsed on
	 * arrival (the parser must be loaded): the newest bytes that parse win, and
	 * only when none do is the key's first trusted embedded font, the one an
	 * eager registration would have parsed, decoded and parsed. Untrusted
	 * embedded candidates are passed over without marking the key attempted.
	 */
	private resolvePending(key: string): void {
		const bytes = this.pendingBytes.get(key);
		const embedded = this.pendingEmbedded.get(key);
		this.dropPending(key);
		if (bytes) {
			for (let i = bytes.length - 1; i >= 0; i -= 1) {
				if (this.parseInto(key, bytes[i]!)) {
					return;
				}
			}
		}
		for (const { font, list } of embedded ?? []) {
			if (!isEmbeddedTypefaceTrusted(list, font.name)) {
				continue;
			}
			this.attempted.add(key);
			const clear = resolveEmbeddedFontClearBytes(font);
			if (clear && clear.length >= 4) {
				// Unparseable fonts (corrupt, or an outline format opentype.js does
				// not support) stay unset so the traced fallback applies.
				this.parseInto(key, clear);
			}
			return;
		}
	}

	private resolveAllPending(): void {
		for (const key of new Set([...this.pendingBytes.keys(), ...this.pendingEmbedded.keys()])) {
			this.resolvePending(key);
		}
	}

	/** Parse the waiting candidates of the keys lookups missed, the only ones a repaint needs. */
	private resolveMissed(): void {
		for (const key of [...this.missed]) {
			if (this.hasPending(key)) {
				this.resolvePending(key);
			}
		}
	}

	private resolveMissedOrLoad(): void {
		if (isOpenTypeFontParserReady()) {
			this.resolveMissed();
		} else {
			this.requestParser();
		}
	}

	private loadParser(): Promise<void> {
		this.loading ??= preloadOpenTypeFontParser().then(
			() => {
				this.loading = undefined;
				this.resolveMissed();
			},
			(error: unknown) => {
				this.loading = undefined;
				throw error;
			},
		);
		return this.loading;
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
		this.loadParser().catch(() => undefined);
	}
}
