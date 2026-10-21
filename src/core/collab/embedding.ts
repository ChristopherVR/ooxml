// The one embedding contract every product editor (Word, Excel, PowerPoint, Visio) takes for
// real-time collaboration, so a host embedding more than one product integrates once: the
// `collaboration` options a host sets, the state an editor reports through `collaborationState`
// and the `collaboration-change` event, the room session an editor opens when the host gives a
// room rather than a session, and the view of a host-owned session that borrowing viewers observe.
// New code for issue #152, generalised from the Excel editor's `XlsxCollaborationOptions`.
import { canBroadcast, createBroadcastTransport } from './broadcast-transport';
import type { ExternalSession, ExternalSnapshot } from './external-provider';
import type { CollaborationRole } from './identity';
import type { ConnectionStatus, ProviderFactory } from './provider';
import { type CollabSession, createCollabSession } from './session';
import { transportProvider } from './transport-provider';
import { isValidRoomId } from './validation';
import { type WebrtcOptions, webrtcProvider } from './webrtc-provider';
import { createWebSocketTransport, roomUrl } from './websocket-transport';

/** Who this window is in the room. */
export interface CollaborationUser {
	/** Display name; each editor falls back to its own author name. */
	name?: string;
	/** Hex colour; defaults to a palette colour derived from the client id. */
	color?: string;
	avatar?: string;
	/** `viewer` makes the editor read-only (advisory: a server must enforce it). */
	role?: CollaborationRole;
}

/**
 * How an editor joins a shared document. One transport is used, in this order: `session`
 * (host-owned, never destroyed by the editor), `provider`, `webrtc`, `serverUrl` (a y-websocket
 * compatible server), else a BroadcastChannel between tabs and windows of this browser. `P` is the product's
 * presence payload, only relevant to a host-made `session`.
 */
export interface CollaborationOptions<P extends object = object> {
	/** Room name: 1-128 letters, digits, hyphens or underscores. Not needed with `session`. */
	roomId?: string;
	/** Base URL of a y-websocket compatible server (`wss://host/collab`); the room is appended. */
	serverUrl?: string;
	/** Sent as the `token` query parameter to `serverUrl`, and as the room password over `webrtc`. */
	authToken?: string;
	/** A provider factory of your own (any transport, or a provider through `adaptYjsProvider`). */
	provider?: ProviderFactory;
	/** Peer-to-peer over the host's `y-webrtc` (see {@link webrtcProvider}); `password` defaults to `authToken`. */
	webrtc?: WebrtcOptions;
	/**
	 * A session the host created, connected and will destroy. Create it with the product's
	 * presence sanitiser so peers validate each other the same way.
	 */
	session?: CollabSession<P>;
	user?: CollaborationUser;
	/**
	 * `join` never seeds an empty room with the local document, `create` may. Products whose
	 * model merges (Excel) seed an empty room either way and ignore it. Default `create`.
	 */
	intent?: 'create' | 'join';
}

/** One person in the room, as title bars and `collaboration-change` list them. */
export interface Collaborator {
	clientId: number;
	name: string;
	color: string;
	self: boolean;
}

/** Reported by every editor's `collaborationState` and `collaboration-change` event. */
export interface CollaborationState {
	active: boolean;
	roomId: string | undefined;
	/** `off` when not sharing; otherwise the provider's connection status. */
	status: ConnectionStatus | 'off';
	/** The first exchange with the room finished. */
	synced: boolean;
	people: Collaborator[];
}

/** The event every editor element dispatches, with a {@link CollaborationState} detail. */
export const COLLABORATION_CHANGE_EVENT = 'collaboration-change';

export const COLLABORATION_OFF_STATE: CollaborationState = Object.freeze({
	active: false,
	roomId: undefined,
	status: 'off',
	synced: false,
	people: [],
});

/**
 * What every editor element exposes, whatever its product. `collaboration` is declarative: while
 * set, the open document is shared; `null` leaves the room.
 */
export interface CollaborativeEditor<P extends object = object> {
	collaboration: CollaborationOptions<P> | null;
	readonly collaborationState: CollaborationState;
	/** Leave the room and clear `collaboration`. A host-owned session stays alive. */
	stopCollaboration(): void;
	/** Restart the connection, keeping the shared document and queued offline changes. */
	reconnectCollaboration(): void;
}

/** The product-specific parts of a room session an editor opens itself. */
export interface RoomSessionDefaults<P extends object> {
	/** The user name when `options.user.name` is empty (the editor's author name). */
	defaultName: string;
	/** BroadcastChannel name prefix, one per product so their rooms never cross. */
	broadcastPrefix: string;
	sanitizePayload?: (raw: Record<string, unknown>) => P | null;
	initialPresence?: P;
}

export const INVALID_ROOM_MESSAGE =
	'Use 1 to 128 letters, digits, hyphens or underscores for the room name.';
export const NO_BROADCAST_MESSAGE = 'This browser cannot share between windows; pass a server URL.';

/**
 * The session for `options`: the host's own (`owned: false`, never destroy it) or a new one this
 * caller owns, connected through `provider`, `serverUrl` or a BroadcastChannel, with pagehide
 * teardown. Throws {@link INVALID_ROOM_MESSAGE} or {@link NO_BROADCAST_MESSAGE}.
 */
export function openCollaborationSession<P extends object>(
	options: CollaborationOptions<P>,
	defaults: RoomSessionDefaults<P>,
): { session: CollabSession<P>; owned: boolean } {
	if (options.session) return { session: options.session, owned: false };
	const roomId = options.roomId ?? '';
	if (!isValidRoomId(roomId)) throw new Error(INVALID_ROOM_MESSAGE);
	let provider = options.provider;
	if (!provider && options.webrtc) {
		const password = options.webrtc.password ?? options.authToken;
		provider = webrtcProvider({ ...options.webrtc, ...(password ? { password } : {}) });
	}
	if (!provider && options.serverUrl)
		provider = transportProvider({
			transport: createWebSocketTransport({
				url: withToken(roomUrl(options.serverUrl, roomId), options.authToken),
			}),
		});
	if (!provider) {
		if (!canBroadcast()) throw new Error(NO_BROADCAST_MESSAGE);
		provider = transportProvider({
			transport: createBroadcastTransport({ roomId, prefix: defaults.broadcastPrefix }),
		});
	}
	const user = options.user ?? {};
	const session = createCollabSession<P>({
		roomId,
		provider,
		user: {
			name: user.name || defaults.defaultName,
			...(user.color ? { color: user.color } : {}),
			...(user.avatar ? { avatar: user.avatar } : {}),
			...(user.role ? { role: user.role } : {}),
		},
		...(defaults.initialPresence ? { initialPresence: defaults.initialPresence } : {}),
		...(defaults.sanitizePayload ? { sanitizePayload: defaults.sanitizePayload } : {}),
		teardown: true,
	});
	return { session, owned: true };
}

function withToken(url: string, token: string | undefined): string {
	if (!token) return url;
	return `${url}${url.includes('?') ? '&' : '?'}token=${encodeURIComponent(token)}`;
}

/** The {@link CollaborationState} of a joined session, this window first. */
export function collaborationStateOf(session: CollabSession<object>): CollaborationState {
	return {
		active: true,
		roomId: session.roomId,
		status: session.status,
		synced: session.synced,
		people: [
			{
				clientId: session.clientId,
				name: session.identity.userName,
				color: session.identity.userColor,
				self: true,
			},
			...session.peers().map((peer) => ({
				clientId: peer.clientId,
				name: peer.userName,
				color: peer.userColor,
				self: false,
			})),
		],
	};
}

/** Calls `listener` whenever the session's {@link CollaborationState} may have changed. */
export function onCollaborationStateChange(
	session: CollabSession<object>,
	listener: () => void,
): () => void {
	const stops = [
		session.on('status', listener),
		session.on('synced', listener),
		session.on('peers', listener),
	];
	return () => {
		for (const stop of stops) stop();
	};
}

/**
 * A {@link CollabSession} as the borrowed {@link ExternalSession} that viewers taking host-owned
 * Yjs resources observe (PowerPoint's `externalSession`). `synced` is the session's write gate:
 * open once the provider synced, or after the grace period for a lone peer on a mesh. The view
 * never connects or destroys the session.
 */
export function toExternalSession(session: CollabSession<object>): ExternalSession {
	// One object per distinct state, so `useSyncExternalStore`-style readers see a stable snapshot.
	let snapshot: ExternalSnapshot = { status: session.status, synced: false };
	return {
		doc: session.doc,
		awareness: session.awareness,
		getSnapshot: (): ExternalSnapshot => {
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
}
