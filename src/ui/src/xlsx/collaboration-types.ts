// The public shape of `<xlsx-editor>` collaboration: the shared `ooxml-core/collab` embedding
// contract every product editor takes, with spreadsheet presence. Types only (and one frozen
// constant), so importing them never loads the Yjs runtime.
import type { CollaborationOptions, CollaborationState, Collaborator } from 'ooxml-core/collab';
import type { XlsxPresence } from 'ooxml-core/xlsx/collab';

/**
 * How to join a shared workbook (see `CollaborationOptions`): `session`, `provider`, `serverUrl`,
 * else a BroadcastChannel. A host-made `session` needs `sanitizePayload: sanitizeXlsxPresence`.
 * The user name defaults to `authorName`. `intent` is ignored: an empty room is always seeded.
 */
export type XlsxCollaborationOptions = CollaborationOptions<XlsxPresence>;

/** One person in the room, as the title bar and `collaboration-change` list them. */
export type XlsxCollaborator = Collaborator;

/** Collaboration state, reported by `collaborationState` and the `collaboration-change` event. */
export type XlsxCollaborationState = CollaborationState;

export const OFF_STATE: XlsxCollaborationState = Object.freeze({
	active: false,
	roomId: undefined,
	status: 'off',
	synced: false,
	people: [],
});
