// Ordering, de-duplication and conflict-classification helpers for the message envelopes that
// authority-based collaboration (docx-viewer's step batches, presence messages) exchanges. Nothing
// here knows about ProseMirror: a message is `{ clientId, sequence | version, fingerprint }`.
// Extracted from docx-viewer `collaboration.ts` / `collaboration-protocol.ts` / `presence.ts`;
// see PROVENANCE.md.
import { isNonNegativeSafeInteger } from './validation.js';

export type OrderingStatus = 'apply' | 'duplicate' | 'stale' | 'out-of-order' | 'invalid';

/**
 * Compare a message's base version with the receiver's current version. Authority-ordered streams
 * must be applied strictly in order: older is `stale` (already seen), newer is `out-of-order`
 * (a gap: the receiver must fetch the missing versions first).
 */
export function classifyVersion(
	messageVersion: number,
	currentVersion: number,
): 'apply' | 'stale' | 'out-of-order' | 'invalid' {
	if (!isNonNegativeSafeInteger(messageVersion) || !isNonNegativeSafeInteger(currentVersion))
		return 'invalid';
	if (messageVersion < currentVersion) return 'stale';
	return messageVersion > currentVersion ? 'out-of-order' : 'apply';
}

/** A Map that forgets its oldest entries beyond `limit`, keeping memory bounded in long sessions. */
export class BoundedMap<K, V> {
	private readonly entries = new Map<K, V>();
	constructor(private readonly limit: number) {
		if (!Number.isSafeInteger(limit) || limit < 1) throw new Error('limit must be positive');
	}
	get size(): number {
		return this.entries.size;
	}
	get(key: K): V | undefined {
		return this.entries.get(key);
	}
	has(key: K): boolean {
		return this.entries.has(key);
	}
	set(key: K, value: V): void {
		this.entries.delete(key);
		this.entries.set(key, value);
		if (this.entries.size > this.limit) {
			const oldest = this.entries.keys().next();
			if (!oldest.done) this.entries.delete(oldest.value);
		}
	}
	delete(key: K): boolean {
		return this.entries.delete(key);
	}
	clear(): void {
		this.entries.clear();
	}
}

/**
 * Idempotency cache for request/ack protocols: a retried message with the same id and the same
 * content is a `duplicate` (replay the earlier result), the same id with different content is
 * `reused` (a client bug or an attack), anything else is `fresh`.
 */
export class IdempotencyCache<V = undefined> {
	private readonly seen: BoundedMap<string, { fingerprint: string; value: V }>;
	constructor(limit = 5_000) {
		this.seen = new BoundedMap(limit);
	}
	check(
		clientId: string,
		messageId: string,
		fingerprint: string,
	): { status: 'fresh' } | { status: 'duplicate'; value: V } | { status: 'reused' } {
		const prior = this.seen.get(`${clientId}\u0000${messageId}`);
		if (!prior) return { status: 'fresh' };
		return prior.fingerprint === fingerprint
			? { status: 'duplicate', value: prior.value }
			: { status: 'reused' };
	}
	remember(clientId: string, messageId: string, fingerprint: string, value: V): void {
		this.seen.set(`${clientId}\u0000${messageId}`, { fingerprint, value });
	}
}

/**
 * Per-peer monotonic sequence tracker for transient streams such as presence: a lower sequence is
 * `stale`, an equal one is a `duplicate` of the same content or `invalid` when the content differs
 * (sequence reuse), and a higher one is accepted and remembered.
 */
export class SequenceTracker {
	private readonly last: BoundedMap<string, { sequence: number; fingerprint: string }>;
	constructor(limit = 200) {
		this.last = new BoundedMap(limit);
	}
	/** Classify without recording; call {@link commit} once the message was actually applied. */
	classify(peerId: string, sequence: number, fingerprint: string): OrderingStatus {
		if (!isNonNegativeSafeInteger(sequence) || sequence < 1) return 'invalid';
		const previous = this.last.get(peerId);
		if (!previous || sequence > previous.sequence) return 'apply';
		if (sequence < previous.sequence) return 'stale';
		return previous.fingerprint === fingerprint ? 'duplicate' : 'invalid';
	}
	commit(peerId: string, sequence: number, fingerprint: string): void {
		this.last.set(peerId, { sequence, fingerprint });
	}
	forget(peerId: string): void {
		this.last.delete(peerId);
	}
}

/** Recursively freeze a JSON-like value so a retained envelope cannot be mutated by a transport. */
export function freezeDeep<T>(value: T): T {
	if (value && typeof value === 'object' && !Object.isFrozen(value)) {
		for (const child of Object.values(value)) freezeDeep(child);
		Object.freeze(value);
	}
	return value;
}
