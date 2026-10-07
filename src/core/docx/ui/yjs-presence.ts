import type { Plugin, EditorState, Transaction } from 'prosemirror-state';
import type { DecorationAttrs } from 'prosemirror-view';
import { yCursorPlugin, yCursorPluginKey } from 'y-prosemirror';
import type { CollabSession } from '../../collab/index.js';

export interface WordYjsPresenceOptions {
	cursorBuilder(user: unknown, clientId: number): HTMLElement;
	selectionBuilder(user: unknown, clientId: number): DecorationAttrs;
}

/** Relative cursor/selection positions over the shared session's awareness transport.
 * The UI owns the DOM builders and localized labels. */
export function wordYjsPresencePlugin(
	session: CollabSession,
	options: WordYjsPresenceOptions,
): Plugin {
	return yCursorPlugin(session.awareness, options);
}

export function refreshWordYjsPresence(state: EditorState): Transaction {
	return state.tr.setMeta(yCursorPluginKey, { awarenessUpdated: true });
}
