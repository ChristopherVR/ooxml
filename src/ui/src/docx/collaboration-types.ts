// The public shape of `<docx-editor>` collaboration: the shared `ooxml-core/collab` embedding
// contract every product editor takes. Types only (and one frozen constant), so importing them
// never loads the Yjs runtime.
import type { CollaborationOptions, CollaborationState, Collaborator } from 'ooxml-core/collab';

/**
 * How to join a shared document (see `CollaborationOptions`): `session`, `provider`, `serverUrl`,
 * else a BroadcastChannel. The user name defaults to `reviewAuthor`. The room name doubles as the
 * Word `documentId`, so every participant must have loaded the same source package. `intent:
 * 'join'` never seeds an empty room; anything else lets this window seed it.
 */
export type DocxCollaborationOptions = CollaborationOptions;

/** One person in the room, as `collaboration-change` lists them. */
export type DocxCollaborator = Collaborator;

/** Collaboration state, reported by `collaborationState` and the `collaboration-change` event. */
export type DocxCollaborationState = CollaborationState;

export const OFF_STATE: DocxCollaborationState = Object.freeze({
	active: false,
	roomId: undefined,
	status: 'off',
	synced: false,
	people: [],
});
