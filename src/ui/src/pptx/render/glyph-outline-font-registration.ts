/**
 * Document lifetime of the embedded fonts a `GlyphOutlineFontCache` holds
 * before they are needed (see `text-warp-outline-font-cache.ts`).
 *
 * The cache only records a deck's embedded fonts when they are registered and
 * decodes one when its `(family, bold, italic)` is first looked up. A deck
 * without envelope WordArt never looks any up, so without a release every
 * recorded font, with its original bytes and the deck's whole font list,
 * would stay in the cache for as long as the cache lives (each binding keeps
 * one module-wide cache). This module holds those records per document:
 *
 *  - {@link EmbeddedFontCandidates} keeps the records per key, in registration
 *    order, and removes one document's records when its last owner releases
 *    it. Registering a list that is already registered adds an owner but no
 *    records, so two viewers showing the same list share them.
 *  - {@link createEmbeddedOutlineFontsSlot} is what a viewer holds: it tracks
 *    the deck it currently shows, releases the previous deck when another
 *    one is set, and releases the current one when the viewer goes away.
 *
 * A document is identified by its font list (`PptxEmbeddedFont[]` from one
 * load), which is also what the trust check is memoised on.
 *
 * @module render/glyph-outline-font-registration
 */
import type { PptxEmbeddedFont } from 'ooxml-core/pptx';

/** An embedded variant recorded for a key, with the list it came from. */
export interface EmbeddedFontCandidate {
	font: PptxEmbeddedFont;
	list: readonly PptxEmbeddedFont[];
}

/** Releases one document's embedded fonts from the cache it was registered with. */
export type ReleaseEmbeddedFonts = () => void;

/** One live registration of a list: the keys it added candidates under, and how many owners hold it. */
interface Registration {
	keys: Set<string>;
	owners: number;
}

/**
 * Embedded font candidates per cache key, grouped by the document (font list)
 * that registered them.
 */
export class EmbeddedFontCandidates {
	private readonly byKey = new Map<string, EmbeddedFontCandidate[]>();
	/** Weak, so a list that is no longer referenced by any candidate is not kept here. */
	private readonly registrations = new WeakMap<readonly PptxEmbeddedFont[], Registration>();

	/**
	 * Record each font of `list` under `keyOf(font)`, after the candidates
	 * already recorded for that key. `keyOf` returns `undefined` for a font
	 * that needs no record. A list that is already registered is left as it is,
	 * so its fonts keep their place, and gains one more owner. Returns this
	 * owner's release: the candidates `list` has here are removed once every
	 * owner has released. Calling a release again does nothing, and a release
	 * from before the list was fully released never touches a later
	 * registration of the same list.
	 */
	add(
		list: readonly PptxEmbeddedFont[],
		keyOf: (font: PptxEmbeddedFont) => string | undefined,
	): ReleaseEmbeddedFonts {
		let registration = this.registrations.get(list);
		if (!registration) {
			registration = { keys: new Set(), owners: 0 };
			this.registrations.set(list, registration);
			for (const font of list) {
				const key = keyOf(font);
				if (key === undefined) {
					continue;
				}
				registration.keys.add(key);
				const candidates = this.byKey.get(key);
				if (candidates) {
					candidates.push({ font, list });
				} else {
					this.byKey.set(key, [{ font, list }]);
				}
			}
		}
		registration.owners += 1;
		const owned = registration;
		let released = false;
		return () => {
			if (released) {
				return;
			}
			released = true;
			this.release(list, owned);
		};
	}

	/** Whether any candidate is recorded for `key`. */
	has(key: string): boolean {
		return this.byKey.has(key);
	}

	/** The keys that have candidates recorded. */
	keys(): IterableIterator<string> {
		return this.byKey.keys();
	}

	/** Remove and return the candidates recorded for `key`, in registration order. */
	take(key: string): EmbeddedFontCandidate[] | undefined {
		const candidates = this.byKey.get(key);
		this.byKey.delete(key);
		return candidates;
	}

	/** Every candidate still recorded, key by key in registration order. */
	*[Symbol.iterator](): IterableIterator<EmbeddedFontCandidate> {
		for (const candidates of this.byKey.values()) {
			yield* candidates;
		}
	}

	private release(list: readonly PptxEmbeddedFont[], registration: Registration): void {
		if (this.registrations.get(list) !== registration) {
			return;
		}
		registration.owners -= 1;
		if (registration.owners > 0) {
			return;
		}
		this.registrations.delete(list);
		for (const key of registration.keys) {
			const candidates = this.byKey.get(key);
			if (!candidates) {
				continue;
			}
			const kept = candidates.filter((candidate) => candidate.list !== list);
			if (kept.length > 0) {
				this.byKey.set(key, kept);
			} else {
				this.byKey.delete(key);
			}
		}
	}
}

/** What {@link createEmbeddedOutlineFontsSlot} needs from a `GlyphOutlineFontCache`. */
export interface EmbeddedOutlineFontsRegistrar {
	registerEmbeddedFonts(fonts: readonly PptxEmbeddedFont[]): ReleaseEmbeddedFonts;
}

/** One viewer's registration of the deck it shows (see {@link createEmbeddedOutlineFontsSlot}). */
export interface EmbeddedOutlineFontsSlot {
	/**
	 * Register `fonts` as the viewer's current deck. Setting another list
	 * releases the previous one first; setting the current list again does
	 * nothing. `null`, `undefined` and an empty list just release.
	 */
	set(fonts: readonly PptxEmbeddedFont[] | null | undefined): void;
	/** Release the current deck, when the viewer is destroyed. Safe to call repeatedly. */
	release(): void;
}

/**
 * The embedded-font registration a viewer keeps with `cache` for the deck it
 * shows: call `set` when a deck is loaded (or its font list is replaced) and
 * `release` when the viewer is torn down.
 */
export function createEmbeddedOutlineFontsSlot(
	cache: EmbeddedOutlineFontsRegistrar,
): EmbeddedOutlineFontsSlot {
	let current: readonly PptxEmbeddedFont[] | undefined;
	let releaseCurrent: ReleaseEmbeddedFonts | undefined;
	const release = (): void => {
		const releasing = releaseCurrent;
		current = undefined;
		releaseCurrent = undefined;
		releasing?.();
	};
	return {
		set(fonts) {
			if (fonts && fonts === current) {
				return;
			}
			release();
			if (!fonts || fonts.length === 0) {
				return;
			}
			current = fonts;
			releaseCurrent = cache.registerEmbeddedFonts(fonts);
		},
		release,
	};
}
