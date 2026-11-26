/**
 * @fileoverview Content fingerprint of a slide, used to decide whether the
 * save pipeline has to re-serialize it at all.
 *
 * ## Why a fingerprint rather than a dirty flag
 *
 * `PptxSlide.isDirty` was designed as an announcement: a mutation path sets it
 * and {@link processSlideForSave} skips everything else. Repo-wide it was set
 * in two places and never to `false`, so the guard was dead and every save
 * re-serialized every slide, flattening inheritance on slides the user never
 * touched.
 *
 * Announcing does not work here. There are >100 slide-mutation sites across
 * five bindings plus the SDK, the AI tool runner and the collaboration
 * reconciler, and every one of them would have to remember. One that forgot
 * would silently DISCARD the user's edit, which is far worse than a redundant
 * write. Worse still, the bindings hand `save()` fresh slide objects on every
 * call (immutable state updates), so nothing identity-based can work either.
 *
 * Observing the value works instead: hash the slide model at load, hash it
 * again at save, and re-serialize when the two differ. A mutation path cannot
 * forget to announce something it does not have to announce, and any field
 * added to `PptxSlide` in future is covered the moment it exists, in the safe
 * direction (unknown change -> looks dirty -> full re-serialization).
 *
 * ## Guarantees
 *
 * - Key order is irrelevant (`{a,b}` and `{b,a}` hash the same), because
 *   immutable spreads reorder keys freely.
 * - An absent key and a key set to `undefined` hash the same, for the same
 *   reason.
 * - `isDirty` is excluded: it is the flag under discussion, not content.
 * - Cycles terminate.
 * - Time is linear in the distinct objects reachable from the slide. Each object (array, typed
 *   array or record) is hashed from a fresh state and enters its parent as a 64-bit digest, so its
 *   contribution depends on its content alone, never on where else it appears. A subtree the
 *   model shares by reference (one parsed XML node reached from many places) is therefore hashed
 *   at most twice: its digest is kept from its second visit on. Keeping it only then leaves an
 *   ordinary, unshared slide no slower than a single walk.
 *
 * @module core/core/runtime/slide-fingerprint
 */
import { at } from '../../../../geometry/indexed';
import type { PptxSlide } from '../../types';

/**
 * Above this length a string is sampled rather than hashed whole. Data URLs
 * for embedded video can be tens of megabytes and a full walk of every one of
 * them on every save would cost more than the re-serialization this exists to
 * avoid. The sample covers the head, the tail and a strided sweep of the
 * middle, so any real content swap (which changes bytes densely, or at least
 * changes the length) is caught.
 */
const FULL_HASH_LIMIT = 1 << 20;

/** Characters sampled from the middle of an over-long string. */
const SAMPLE_BUDGET = 1 << 15;

/**
 * Slide-level keys deliberately left out of the hash.
 *
 * - `isDirty` is the flag this whole mechanism replaces, not content.
 * - `slideNumber` is the slide's POSITION in the deck. Nothing in a
 *   `ppt/slides/slideN.xml` part depends on it (deck order lives in
 *   `p:sldIdLst`), and the save pipeline renumbers it itself on every save
 *   (`PptxPresentationSlidesReconciler`). Hashing it would mark every slide
 *   after a deleted or reordered one dirty, which is precisely the 99-slide
 *   rewrite this exists to prevent.
 *
 * Nothing else may join this list without the same two-part argument: the
 * save pipeline must not read it, and something outside the user's control
 * must be changing it.
 */
const POSITION_ONLY_SLIDE_KEYS: readonly string[] = ['isDirty', 'slideNumber'];

/**
 * Streaming 64-bit string hash (two decorrelated 32-bit lanes, cyrb-style).
 * Not cryptographic: it only has to make an accidental collision between two
 * different slide states implausible.
 */
const SEED1 = 0xdeadbeef;
const SEED2 = 0x41c6ce57;

class SlideHasher {
	public h1 = SEED1;

	public h2 = SEED2;

	public updateChar(character: string): void {
		const code = character.charCodeAt(0);
		this.h1 = Math.imul(this.h1 ^ code, 2654435761);
		this.h2 = Math.imul(this.h2 ^ code, 1597334677);
	}

	public update(text: string, stride = 1): void {
		let h1 = this.h1;
		let h2 = this.h2;
		for (let index = 0; index < text.length; index += stride) {
			const code = text.charCodeAt(index);
			h1 = Math.imul(h1 ^ code, 2654435761);
			h2 = Math.imul(h2 ^ code, 1597334677);
		}
		this.h1 = h1;
		this.h2 = h2;
	}

	/** Hashes byte codes directly, without constructing a temporary string. */
	public updateBytes(bytes: Uint8Array, stride: number): void {
		let h1 = this.h1;
		let h2 = this.h2;
		for (let index = 0; index < bytes.length; index += stride) {
			const code = bytes[index] ?? 0;
			h1 = Math.imul(h1 ^ code, 2654435761);
			h2 = Math.imul(h2 ^ code, 1597334677);
		}
		this.h1 = h1;
		this.h2 = h2;
	}

	/** Folds a nested object's digest into this state. */
	public mix(d1: number, d2: number): void {
		// Multipliers other than `update`'s, so a digest never reads as a character.
		this.h1 = Math.imul(this.h1 ^ d1, 0x85ebca6b);
		this.h2 = Math.imul(this.h2 ^ d2, 0xc2b2ae35);
	}

	/** The two finalized 32-bit lanes of the current state. */
	public final1(): number {
		return (
			Math.imul(this.h1 ^ (this.h1 >>> 16), 2246822507) ^
			Math.imul(this.h2 ^ (this.h2 >>> 13), 3266489909)
		);
	}

	public final2(): number {
		return (
			Math.imul(this.h2 ^ (this.h2 >>> 16), 2246822507) ^
			Math.imul(this.h1 ^ (this.h1 >>> 13), 3266489909)
		);
	}

	public digest(): string {
		const h1 = this.final1();
		const h2 = this.final2();
		return `${(h2 >>> 0).toString(16).padStart(8, '0')}${(h1 >>> 0).toString(16).padStart(8, '0')}`;
	}
}

/** An object being hashed, one hashed once (digest not kept), or a kept digest. */
const ON_PATH = 1;
const SEEN_ONCE = 2;
type Visit = typeof ON_PATH | typeof SEEN_ONCE | { d1: number; d2: number };

function hashString(hasher: SlideHasher, text: string): void {
	hasher.update(`s${text.length}:`);
	if (text.length <= FULL_HASH_LIMIT) {
		hasher.update(text);
		return;
	}
	const window = SAMPLE_BUDGET >> 2;
	hasher.update(text.slice(0, window));
	hasher.update(text.slice(text.length - window));
	const stride = Math.ceil(text.length / (SAMPLE_BUDGET >> 1));
	hasher.update(text, stride);
}

function hashValue(hasher: SlideHasher, value: unknown, visits: Map<object, Visit>): void {
	if (value === undefined) {
		hasher.updateChar('u');
		return;
	}
	if (value === null) {
		hasher.updateChar('n');
		return;
	}
	switch (typeof value) {
		case 'string':
			hashString(hasher, value);
			return;
		case 'number':
			hasher.update(`d${value}`);
			return;
		case 'boolean':
			hasher.update(value ? 'b1' : 'b0');
			return;
		case 'bigint':
			hasher.update(`g${value}`);
			return;
		case 'function':
			// A callback on the model is not file content; its identity changes
			// on every render in every binding, so hashing it would make every
			// slide look dirty forever.
			hasher.updateChar('f');
			return;
		case 'symbol':
			hasher.updateChar('y');
			return;
		default:
			break;
	}

	const object = value as object;
	const visit = visits.get(object);
	if (visit === ON_PATH) {
		hasher.updateChar('c');
		return;
	}
	if (visit !== undefined && visit !== SEEN_ONCE) {
		hasher.mix(visit.d1, visit.d2);
		return;
	}
	visits.set(object, ON_PATH);
	const outer1 = hasher.h1;
	const outer2 = hasher.h2;
	hasher.h1 = SEED1;
	hasher.h2 = SEED2;
	hashContent(hasher, object, visits);
	const d1 = hasher.final1();
	const d2 = hasher.final2();
	hasher.h1 = outer1;
	hasher.h2 = outer2;
	hasher.mix(d1, d2);
	visits.set(object, visit === SEEN_ONCE ? { d1, d2 } : SEEN_ONCE);
}

function hashContent(hasher: SlideHasher, object: object, visits: Map<object, Visit>): void {
	if (Array.isArray(object)) {
		hasher.update(`a${object.length}[`);
		for (const item of object) {
			hashValue(hasher, item, visits);
			hasher.updateChar(',');
		}
		hasher.updateChar(']');
	} else if (ArrayBuffer.isView(object)) {
		const bytes = new Uint8Array(
			(object as ArrayBufferView).buffer,
			(object as ArrayBufferView).byteOffset,
			(object as ArrayBufferView).byteLength,
		);
		hasher.update(`v${bytes.length}:`);
		const stride = bytes.length > FULL_HASH_LIMIT ? Math.ceil(bytes.length / SAMPLE_BUDGET) : 1;
		hasher.updateBytes(bytes, stride);
	} else {
		const record = object as Record<string, unknown>;
		const keys = Object.keys(record);
		if (keys.length === 2) {
			// Object.keys returns a dense array. Avoid a general sort for a node pair.
			const first = at(keys, 0);
			const second = at(keys, 1);
			if (first > second) {
				keys[0] = second;
				keys[1] = first;
			}
		} else if (keys.length > 2) keys.sort();
		hasher.update('o{');
		for (const key of keys) {
			const entry = record[key];
			// An absent key and an explicitly-undefined key are the same slide.
			if (entry === undefined) {
				continue;
			}
			hasher.update(key);
			hasher.updateChar('=');
			hashValue(hasher, entry, visits);
			hasher.updateChar(';');
		}
		hasher.updateChar('}');
	}
}

/**
 * Stable content hash of everything the save pipeline can read off a slide.
 *
 * Two slides with the same fingerprint serialize to the same slide part; two
 * slides that differ anywhere at all, including in a field this module has
 * never heard of, get different fingerprints.
 */
export function fingerprintSlide(slide: PptxSlide): string {
	const hasher = new SlideHasher();
	try {
		const content: Record<string, unknown> = { ...slide };
		for (const key of POSITION_ONLY_SLIDE_KEYS) {
			delete content[key];
		}
		hashValue(hasher, content, new Map<object, Visit>());
	} catch {
		// The walk is recursive, so pathologically deep attacker-supplied XML can
		// exhaust the stack, and a hostile getter on the model can throw. Either
		// way the answer is "cannot prove this slide is unchanged": return a value
		// nothing will ever match, so the slide is re-serialized in full.
		return `unhashable:${UNHASHABLE_COUNTER++}`;
	}
	return hasher.digest();
}

/** Makes every failed fingerprint distinct, so it can never match a baseline. */
let UNHASHABLE_COUNTER = 0;

/** Record `slides` as the state now sitting in the archive. */
export function recordSlideFingerprints(
	fingerprints: Map<string, string>,
	slides: readonly PptxSlide[],
): void {
	for (const slide of slides) {
		fingerprints.set(slide.id, fingerprintSlide(slide));
	}
}

/**
 * True when this slide is byte-for-byte already in the archive.
 *
 * A slide with no recorded baseline (created this session, or arriving after a
 * `dispose()`) is never treated as unchanged: there is nothing in the archive
 * to fall back on, so it must be written.
 */
export function slideMatchesFingerprint(
	fingerprints: Map<string, string>,
	slide: PptxSlide,
): boolean {
	const baseline = fingerprints.get(slide.id);
	return baseline !== undefined && baseline === fingerprintSlide(slide);
}
