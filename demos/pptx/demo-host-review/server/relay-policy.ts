// What a connection may send to a room, decided from the y-websocket message alone. This is the
// server-side half of the viewer's `role`: `viewer` only makes the editor read-only in the browser,
// so the relay drops document updates from read-only connections whatever the client claims.
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';

export const MESSAGE_SYNC = 0;
export const MESSAGE_AWARENESS = 1;
export const MESSAGE_QUERY_AWARENESS = 3;

const SYNC_STEP_1 = 0;
const SYNC_STEP_2 = 1;
const SYNC_UPDATE = 2;

export type MessageKind =
	| 'sync-step1'
	| 'sync-step2'
	| 'sync-update'
	| 'awareness'
	| 'query-awareness'
	| 'invalid';

export function messageKind(bytes: Uint8Array): MessageKind {
	try {
		const decoder = decoding.createDecoder(bytes);
		const type = decoding.readVarUint(decoder);
		if (type === MESSAGE_AWARENESS) return 'awareness';
		if (type === MESSAGE_QUERY_AWARENESS) return 'query-awareness';
		if (type !== MESSAGE_SYNC) return 'invalid';
		const step = decoding.readVarUint(decoder);
		if (step === SYNC_STEP_1) return 'sync-step1';
		if (step === SYNC_STEP_2) return 'sync-step2';
		if (step === SYNC_UPDATE) return 'sync-update';
		return 'invalid';
	} catch {
		return 'invalid';
	}
}

/**
 * `true` when the relay should process the message. Read-only connections may ask for the
 * document (step 1) and share presence, never change the document (step 2 or an update).
 */
export function mayProcess(kind: MessageKind, writable: boolean): boolean {
	if (kind === 'invalid') return false;
	if (kind === 'sync-step2' || kind === 'sync-update') return writable;
	return true;
}

/**
 * Keep only the awareness entries a connection may speak for: client ids it already owns, or ids
 * nobody owns yet (which it then claims), so one participant cannot clear or impersonate another's
 * presence. Returns the re-encoded update, or null when nothing is left.
 */
export function ownAwareness<C>(
	update: Uint8Array,
	connection: C,
	owners: Map<number, C>,
	maxOwned = 4,
): Uint8Array | null {
	const decoder = decoding.createDecoder(update);
	const count = decoding.readVarUint(decoder);
	const kept: { id: number; clock: number; state: string }[] = [];
	for (let i = 0; i < count; i++) {
		const id = decoding.readVarUint(decoder);
		const clock = decoding.readVarUint(decoder);
		const state = decoding.readVarString(decoder);
		const owner = owners.get(id);
		if (owner !== undefined && owner !== connection) continue;
		if (owner === undefined) {
			const owned = [...owners.values()].filter((value) => value === connection).length;
			if (owned >= maxOwned) continue;
			owners.set(id, connection);
		}
		kept.push({ id, clock, state });
	}
	if (kept.length === 0) return null;
	const encoder = encoding.createEncoder();
	encoding.writeVarUint(encoder, kept.length);
	for (const { id, clock, state } of kept) {
		encoding.writeVarUint(encoder, id);
		encoding.writeVarUint(encoder, clock);
		encoding.writeVarString(encoder, state);
	}
	return encoding.toUint8Array(encoder);
}
