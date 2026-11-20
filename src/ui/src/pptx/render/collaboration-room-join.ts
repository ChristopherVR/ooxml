/**
 * When a collaboration room's slides replace the deck the user had open.
 *
 * The first non-empty room content a session adopts, before this editor wrote
 * its own slides to the room, is a join: the room's deck replaces the open one,
 * just as opening a file would. After either has happened, room changes are
 * ordinary edits. A binding reports the join as a presentation load, which is
 * what drops the assistant's staged proposals for the deck that was replaced.
 */
export interface RoomJoinTracker {
	/** A session attached to a room: its next adoption may be the join. */
	attached(): void;
	/** This editor wrote its slides to the room, so the room holds its deck. */
	wrote(): void;
	/** Non-empty room slides were adopted. Returns true when they replaced the open deck. */
	adopted(): boolean;
}

export function createRoomJoinTracker(): RoomJoinTracker {
	let pending = false;
	return {
		attached() {
			pending = true;
		},
		wrote() {
			pending = false;
		},
		adopted() {
			const joined = pending;
			pending = false;
			return joined;
		},
	};
}
