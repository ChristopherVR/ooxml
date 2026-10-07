import type { CollabSession } from 'ooxml-core/collab';
import { sanitizeUserName, sanitizeColor } from 'ooxml-core/collab';
import { wordYjsPresencePlugin } from 'ooxml-core/docx/ui';
import { peerCursor, updatePeerCursor } from './presence-decorations';
import { Plugin } from 'prosemirror-state';
import { translateTemplate, type EditorLocale } from './localization';

export function yjsPresenceProfile(raw: unknown): { name: string; color: string } {
	const value = raw && typeof raw === 'object' ? (raw as Record<string, unknown>) : {};
	return { name: sanitizeUserName(value.name, 'Collaborator'), color: sanitizeColor(value.color) };
}

export function yjsPresence(session: CollabSession, locale: () => EditorLocale) {
	const cursors = new Map<number, HTMLElement>();
	const plugin = wordYjsPresencePlugin(session, {
		cursorBuilder: (user, clientId) => {
			const cursor = peerCursor(yjsPresenceProfile(user), locale());
			cursors.set(clientId, cursor);
			return cursor;
		},
		selectionBuilder: (user) => {
			const peer = yjsPresenceProfile(user);
			return {
				class: 'dve-peer-selection',
				style: `background-color: ${peer.color}33; box-shadow: inset 0 -2px ${peer.color}`,
				'aria-label': translateTemplate(locale(), 'presence.selection', { name: peer.name }),
			};
		},
	});
	return new Plugin({
		...plugin.spec,
		view: (view) => {
			const binding = plugin.spec.view?.(view);
			const refresh = () => {
				for (const [clientId, cursor] of cursors) {
					const state = session.awareness.getStates().get(clientId);
					if (state) updatePeerCursor(cursor, yjsPresenceProfile(state.user), locale());
					else cursors.delete(clientId);
				}
			};
			session.awareness.on('change', refresh);
			return {
				update: (next, previous) => {
					binding?.update?.(next, previous);
					refresh();
				},
				destroy: () => {
					session.awareness.off('change', refresh);
					binding?.destroy?.();
					cursors.clear();
				},
			};
		},
	});
}
