import type { CollabSession } from 'ooxml-core/collab';

import type { CollaborationConfig } from '../types';
import { externalSessionFromCollabSession, isSameBrowserRoom } from './collaboration-options';

/** The default author name when the host sets none. */
export const DEFAULT_OWNED_USER_NAME = 'PowerPoint user';
/** BroadcastChannel name prefix: one per product so their rooms never cross. */
export const POWERPOINT_BROADCAST_PREFIX = 'ooxml-pptx:';

/**
 * Whether PowerPoint's own y-websocket / y-webrtc path cannot serve `config`, so the viewer opens
 * a session of its own. The shared contract order is `session` > `provider` > `webrtc` > `serverUrl` >
 * BroadcastChannel: a host session (`externalSession`) is always borrowed, a `provider` always
 * or `webrtc` wins over `serverUrl`, a `serverUrl` (websocket) or `transport: 'webrtc'` stays on the built-in
 * provider path, and shared-contract input with none of those shares between windows of this
 * browser. A legacy config whose `serverUrl` is blank keeps its serverless WebRTC meaning.
 */
export function needsOwnedSession(config: CollaborationConfig): boolean {
	if (config.externalSession) {
		return false;
	}
	if (config.provider || config.webrtc) {
		return true;
	}
	return isSameBrowserRoom(config);
}

/** A pending or finished owned session, released exactly once by the binding that started it. */
export interface OwnedCollaboration {
	/**
	 * The config with the owned session as its `externalSession`, or `null` when
	 * {@link OwnedCollaboration.release} ran first (the session is then already destroyed).
	 * Rejects with the room or transport error when no session can be opened.
	 */
	readonly ready: Promise<CollaborationConfig | null>;
	/** Destroy the owned session, now or as soon as the lazy runtime import settles. */
	release: () => void;
}

/**
 * Keep the PowerPoint presence payload (slide, cursor, selection) when the session's own
 * publisher heartbeats the identity-only `presence` field of the same awareness, so remote
 * cursors do not flicker back to the origin between PowerPoint's own publishes.
 */
function keepPowerPointPresence(awareness: CollabSession<object>['awareness']): void {
	const write = awareness.setLocalStateField.bind(awareness);
	let latest: Record<string, unknown> | undefined;
	awareness.setLocalStateField = (field: string, value: unknown): void => {
		if (field === 'presence' && value && typeof value === 'object') {
			const next = value as Record<string, unknown>;
			if ('activeSlideIndex' in next) {
				latest = next;
			} else if (latest) {
				write(field, { ...latest, ...next });
				return;
			}
		}
		write(field, value);
	};
}

/**
 * Open the session `config` asks for through the shared `openCollaborationSession` (a `provider`
 * or, without one, a BroadcastChannel between windows of this browser). `ooxml-core/collab`
 * (and with it the Yjs runtime the bindings already share) is imported only now. The session is
 * presented through the borrowed-session path, so readiness, `sessionIntent: 'join'` and role
 * handling are the ones a host session gets; only the lifetime differs: `release` destroys it.
 */
export function openOwnedCollaboration(config: CollaborationConfig): OwnedCollaboration {
	let released = false;
	let session: CollabSession<object> | undefined;
	const destroy = (): void => {
		const current = session;
		session = undefined;
		current?.destroy();
	};
	const ready = (async (): Promise<CollaborationConfig | null> => {
		const { openCollaborationSession } = await import('ooxml-core/collab');
		if (released) {
			return null;
		}
		const opened = openCollaborationSession(
			{
				roomId: config.roomId,
				...(config.provider ? { provider: config.provider } : {}),
				...(config.webrtc ? { webrtc: config.webrtc } : {}),
				...(config.serverUrl ? { serverUrl: config.serverUrl } : {}),
				...(config.authToken ? { authToken: config.authToken } : {}),
				user: {
					...(config.userName ? { name: config.userName } : {}),
					...(config.userColor ? { color: config.userColor } : {}),
					...(config.userAvatar ? { avatar: config.userAvatar } : {}),
					...(config.role ? { role: config.role } : {}),
				},
			},
			{
				defaultName: config.userName || DEFAULT_OWNED_USER_NAME,
				broadcastPrefix: POWERPOINT_BROADCAST_PREFIX,
			},
		);
		session = opened.session;
		keepPowerPointPresence(opened.session.awareness);
		return {
			...config,
			userName: config.userName || opened.session.identity.userName,
			userColor: config.userColor ?? opened.session.identity.userColor,
			externalSession: externalSessionFromCollabSession(opened.session),
		};
	})();
	return {
		ready,
		release: () => {
			released = true;
			destroy();
		},
	};
}
