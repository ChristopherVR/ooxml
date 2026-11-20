/**
 * When a collaboration room's slides replace the deck the user had open, and
 * when the open deck may be written to the room.
 *
 * The first non-empty room content a session adopts, before this editor wrote
 * its own slides to the room, is a join: the room's deck replaces the open one,
 * just as opening a file would. After either has happened, room changes are
 * ordinary edits. A binding reports the join as a presentation load, which is
 * what drops the assistant's staged proposals for the deck that was replaced.
 *
 * A session that joins someone else's room (`sessionIntent: 'join'`) opened
 * with a placeholder deck, so it never seeds the room with it: it writes only
 * after it adopted the room's slides, or after the user opened a deck of their
 * own. Otherwise a joiner that reached an empty room before the sharer's deck
 * did would write its placeholder, and the sharer's own deck, loading after
 * it, would be replaced by the room's copy of the placeholder.
 */
import type { CollaborationSessionIntent } from '../types';
import type { CollabLoadOrigin } from './collaboration-load-origin';

export interface RoomJoinTracker {
	/** A session attached to a room: its next adoption may be the join. */
	attached(intent?: CollaborationSessionIntent): void;
	/** This editor wrote its slides to the room, so the room holds its deck. */
	wrote(): void;
	/** Non-empty room slides were adopted. Returns true when they replaced the open deck. */
	adopted(): boolean;
	/** A deck finished loading; one the user opened is theirs to publish, even into an empty room. */
	loaded(origin: CollabLoadOrigin): void;
	/** Whether the open deck may be written to the room: false while a join waits for its slides. */
	mayWrite(): boolean;
}

export function createRoomJoinTracker(): RoomJoinTracker {
	let pending = false;
	let awaitingRoom = false;
	return {
		attached(intent) {
			pending = true;
			awaitingRoom = intent === 'join';
		},
		wrote() {
			pending = false;
		},
		adopted() {
			const joined = pending;
			pending = false;
			awaitingRoom = false;
			return joined;
		},
		loaded(origin) {
			if (origin === 'user') {
				awaitingRoom = false;
			}
		},
		mayWrite() {
			return !awaitingRoom;
		},
	};
}
