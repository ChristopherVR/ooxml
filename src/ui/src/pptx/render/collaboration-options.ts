import type { CollaborationOptions, CollabSession } from 'ooxml-core/collab';

import type { CollaborationConfig } from '../types';
import type {
	ExternalCollaborationSession,
	ExternalCollaborationSnapshot,
} from './collaboration-external-session';

/**
 * What a host may set as `collaboration`: PowerPoint's own {@link CollaborationConfig}, or the
 * shared `CollaborationOptions` contract Word, Excel and Visio take (`session`, `user`,
 * `intent`). `provider` is not part of the PowerPoint input: its transports are built in, so a
 * host with its own transport creates a session (`openCollaborationSession` from
 * `ooxml-core/collab`) and passes it as `session`.
 */
export type CollaborationInput =
	| CollaborationConfig
	| (Omit<CollaborationOptions, 'provider'> & { provider?: never });

const externalSessions = new WeakMap<object, ExternalCollaborationSession>();
const resolved = new WeakMap<object, CollaborationConfig>();

/**
 * A host `CollabSession` as the borrowed session the bindings observe. One object per session
 * (the bindings require a stable reference); the snapshot object only changes with its values.
 * Written against the session's public surface so no Yjs runtime is imported here.
 */
export function externalSessionFromCollabSession(
	session: CollabSession<object>,
): ExternalCollaborationSession {
	const known = externalSessions.get(session);
	if (known) return known;
	let snapshot: ExternalCollaborationSnapshot = { status: session.status, synced: false };
	const view: ExternalCollaborationSession = {
		doc: session.doc,
		awareness: session.awareness as unknown as ExternalCollaborationSession['awareness'],
		getSnapshot: () => {
			const status = session.status;
			const synced = session.synced || session.gate.isOpen();
			if (snapshot.status !== status || snapshot.synced !== synced) snapshot = { status, synced };
			return snapshot;
		},
		subscribe: (listener) => {
			const stops = [
				session.on('status', listener),
				session.on('synced', listener),
				session.on('ready', listener),
			];
			return () => {
				for (const stop of stops) stop();
			};
		},
	};
	externalSessions.set(session, view);
	return view;
}

/**
 * Normalise the `collaboration` a host passed into a {@link CollaborationConfig}. Every binding
 * calls this at the one place it reads the prop. A legacy config returns as is (same object);
 * shared-contract input maps `session` to a stable `externalSession`, `user` to the user fields
 * (explicit legacy fields win) and `intent` to `sessionIntent`. A `provider` without a `session`
 * is unsupported and resolves to an empty room id, which every binding reports as the existing
 * `error` status.
 */
export function resolveCollaborationConfig(
	input: CollaborationInput | null | undefined,
): CollaborationConfig | undefined {
	if (!input) return undefined;
	const raw = input as Partial<CollaborationConfig> & {
		session?: CollabSession<object>;
		user?: CollaborationOptions['user'];
		intent?: 'create' | 'join';
		provider?: unknown;
	};
	if (!raw.session && !raw.user && !raw.intent && !raw.provider) {
		return input as CollaborationConfig;
	}
	const cached = resolved.get(input);
	if (cached) return cached;
	const { session, user, intent, provider, ...legacy } = raw;
	const unsupported = Boolean(provider) && !session && !legacy.externalSession;
	const config: CollaborationConfig = {
		...legacy,
		roomId: unsupported ? '' : (legacy.roomId ?? session?.roomId ?? ''),
		serverUrl: legacy.serverUrl ?? '',
		userName: legacy.userName ?? user?.name ?? session?.identity.userName ?? '',
	};
	const userColor = legacy.userColor ?? user?.color ?? session?.identity.userColor;
	if (userColor) config.userColor = userColor;
	const userAvatar = legacy.userAvatar ?? user?.avatar;
	if (userAvatar) config.userAvatar = userAvatar;
	const role = legacy.role ?? user?.role;
	if (role) config.role = role;
	const sessionIntent = legacy.sessionIntent ?? intent;
	if (sessionIntent) config.sessionIntent = sessionIntent;
	if (session && !legacy.externalSession) {
		config.externalSession = externalSessionFromCollabSession(session);
	}
	resolved.set(input, config);
	return config;
}
