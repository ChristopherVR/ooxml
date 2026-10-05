// Update encoding and decoding, snapshots and defensive application of untrusted updates.
// New code written for the collab area; the pptx and docx viewers each hand-rolled y-websocket usage
// and never validated what a peer sent.
import * as Y from 'yjs';

/** Default upper bound for one inbound update or snapshot (32 MiB). */
export const MAX_UPDATE_BYTES = 32 * 1024 * 1024;

export type ApplyResult = { ok: true } | { ok: false; reason: string };

export interface ApplyOptions {
	/** Reject payloads larger than this. Default {@link MAX_UPDATE_BYTES}. */
	maxBytes?: number;
	/** Structurally decode the update before touching the document. Default true. */
	validate?: boolean;
}

/** Encode the whole document state as one update (usable as a persisted snapshot). */
export function encodeSnapshot(doc: Y.Doc): Uint8Array {
	return Y.encodeStateAsUpdate(doc);
}

/** Encode what `doc` has that a peer with state vector `peerStateVector` lacks. */
export function encodeDiff(doc: Y.Doc, peerStateVector?: Uint8Array): Uint8Array {
	return Y.encodeStateAsUpdate(doc, peerStateVector);
}

export function encodeStateVector(doc: Y.Doc): Uint8Array {
	return Y.encodeStateVector(doc);
}

/**
 * Check that `update` is a well-formed Yjs update without applying it. Decoding happens before any
 * integration, so a payload that fails here can never half-apply (a malformed delete set would
 * otherwise throw only after structs were integrated).
 */
export function validateUpdate(update: unknown, maxBytes: number = MAX_UPDATE_BYTES): ApplyResult {
	if (!(update instanceof Uint8Array)) return { ok: false, reason: 'update must be a Uint8Array' };
	if (update.byteLength === 0) return { ok: false, reason: 'update is empty' };
	if (update.byteLength > maxBytes)
		return { ok: false, reason: `update exceeds the ${maxBytes} byte limit` };
	try {
		Y.decodeUpdate(update);
		return { ok: true };
	} catch (cause) {
		return {
			ok: false,
			reason: `malformed update: ${cause instanceof Error ? cause.message : String(cause)}`,
		};
	}
}

/** Apply a possibly hostile update; never throws, reports why a payload was rejected. */
export function applyUpdateSafe(
	doc: Y.Doc,
	update: unknown,
	origin?: unknown,
	options: ApplyOptions = {},
): ApplyResult {
	const checked =
		options.validate === false
			? update instanceof Uint8Array && update.byteLength <= (options.maxBytes ?? MAX_UPDATE_BYTES)
				? ({ ok: true } as const)
				: ({ ok: false, reason: 'invalid or oversized update' } as const)
			: validateUpdate(update, options.maxBytes);
	if (!checked.ok) return checked;
	try {
		Y.applyUpdate(doc, update as Uint8Array, origin);
		return { ok: true };
	} catch (cause) {
		return {
			ok: false,
			reason: `could not apply update: ${cause instanceof Error ? cause.message : String(cause)}`,
		};
	}
}

/** Build a fresh document from a snapshot; throws when the snapshot is invalid. */
export function restoreSnapshot(snapshot: Uint8Array, options: ApplyOptions = {}): Y.Doc {
	const doc = new Y.Doc();
	const result = applyUpdateSafe(doc, snapshot, 'snapshot', options);
	if (!result.ok) {
		doc.destroy();
		throw new Error(result.reason);
	}
	return doc;
}

/** Merge several updates into one without a document (compaction of an update log). */
export function mergeUpdates(updates: readonly Uint8Array[]): Uint8Array {
	return Y.mergeUpdates([...updates]);
}

const CHUNK = 0x8000;

/** Base64 for transports that carry text (JSON, postMessage); works without Buffer or DOM. */
export function toBase64(bytes: Uint8Array): string {
	let binary = '';
	for (let i = 0; i < bytes.length; i += CHUNK)
		binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
	return btoa(binary);
}

/** Inverse of {@link toBase64}; returns `null` for input that is not valid base64. */
export function fromBase64(text: string): Uint8Array | null {
	try {
		const binary = atob(text);
		const bytes = new Uint8Array(binary.length);
		for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
		return bytes;
	} catch {
		return null;
	}
}
