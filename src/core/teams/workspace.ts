// The workspace ties the pieces together for one user: a collab session (chat state plus presence
// over a sync transport chosen from the server config), the chat store, and call sessions with
// signaling chosen the same way. This is what a viewer binding instantiates; everything above it is
// UI. DOM-free; browser classes are injectable. New code.
import type * as Y from 'yjs';
import { createBroadcastTransport } from '../collab/broadcast-transport.js';
import type { BroadcastChannelConstructor } from '../collab/broadcast-transport.js';
import { createClientId } from '../collab/identity.js';
import type { Transport } from '../collab/provider.js';
import { type CollabSession, createCollabSession } from '../collab/session.js';
import { transportProvider } from '../collab/transport-provider.js';
import { createWebSocketTransport } from '../collab/websocket-transport.js';
import type { WebSocketConstructor } from '../collab/websocket-transport.js';
import { type CallSession, type MediaDevicesLike, createCallSession } from './call.js';
import { type ChatStore, createChatStore } from './chat.js';
import { createTabStore, type TabStore } from './tabs.js';
import type { PeerConnectionFactory, StreamFactory } from './peer.js';
import { type TeamsServerConfig, endpointUrl, localServerConfig } from './server-config.js';
import {
	type SignalingChannel,
	createBroadcastSignaling,
	createWebSocketSignaling,
} from './signaling.js';

export type Availability = 'available' | 'busy' | 'away';

/** What each person publishes with their presence. */
export interface TeamsPresence {
	channelId?: string;
	/** Channel the person is typing in right now. */
	typingIn?: string;
	callRoom?: string;
	availability?: Availability;
}

export interface WorkspaceOptions {
	/** Room id for the shared state: 1-128 alphanumeric, `-` or `_`. */
	workspaceId: string;
	user: { id?: string; name: string; color?: string };
	config?: TeamsServerConfig;
	/** Bring your own Y.Doc, for example one restored from a local snapshot. */
	doc?: Y.Doc;
	WebSocket?: WebSocketConstructor;
	BroadcastChannel?: BroadcastChannelConstructor;
	mediaDevices?: MediaDevicesLike;
	createPeerConnection?: PeerConnectionFactory;
	createStream?: StreamFactory;
}

export interface TeamsWorkspace {
	readonly user: { id: string; name: string };
	readonly session: CollabSession<TeamsPresence>;
	readonly chat: ChatStore;
	readonly tabs: TabStore;
	readonly config: TeamsServerConfig;
	/** A call for a channel; call `join()` on it. One call per channel per workspace. */
	call: (channelId: string) => CallSession;
	setActiveChannel: (channelId: string | undefined) => void;
	setTyping: (channelId: string | undefined) => void;
	setAvailability: (value: Availability) => void;
	destroy: () => void;
}

const AVAILABILITY: readonly Availability[] = ['available', 'busy', 'away'];
const safeId = (v: unknown): string | undefined =>
	typeof v === 'string' && /^[\w-]{1,64}$/u.test(v) ? v : undefined;

function sanitizePresence(raw: Record<string, unknown>): TeamsPresence {
	const out: TeamsPresence = {};
	const channelId = safeId(raw.channelId);
	const typingIn = safeId(raw.typingIn);
	const callRoom =
		typeof raw.callRoom === 'string' && /^[\w-]{1,200}$/u.test(raw.callRoom)
			? raw.callRoom
			: undefined;
	if (channelId) out.channelId = channelId;
	if (typingIn) out.typingIn = typingIn;
	if (callRoom) out.callRoom = callRoom;
	out.availability = AVAILABILITY.includes(raw.availability as Availability)
		? (raw.availability as Availability)
		: 'available';
	return out;
}

export function createTeamsWorkspace(options: WorkspaceOptions): TeamsWorkspace {
	const config = options.config ?? localServerConfig();
	const user = { id: options.user.id ?? createClientId(), name: options.user.name };
	const wsCtor = options.WebSocket;
	const bcCtor = options.BroadcastChannel;

	const syncTransport = (roomId: string): Transport =>
		config.mode === 'server' && config.syncUrl
			? createWebSocketTransport({
					url: endpointUrl(config.syncUrl, roomId, config.token),
					...(wsCtor ? { WebSocket: wsCtor } : {}),
				})
			: createBroadcastTransport({ roomId, ...(bcCtor ? { BroadcastChannel: bcCtor } : {}) });

	const signaling = (roomId: string): SignalingChannel =>
		config.mode === 'server' && config.signalingUrl
			? createWebSocketSignaling({
					url: endpointUrl(config.signalingUrl, roomId, config.token),
					...(wsCtor ? { WebSocket: wsCtor } : {}),
				})
			: createBroadcastSignaling({ roomId, ...(bcCtor ? { BroadcastChannel: bcCtor } : {}) });

	let chat!: ChatStore;
	const session = createCollabSession<TeamsPresence>({
		roomId: options.workspaceId,
		...(options.doc ? { doc: options.doc } : {}),
		provider: transportProvider({ transport: syncTransport(options.workspaceId) }),
		user: { name: user.name, ...(options.user.color ? { color: options.user.color } : {}) },
		initialPresence: { availability: 'available' },
		sanitizePayload: sanitizePresence,
		teardown: true,
		// Any writer may seed the default channel; the fixed id makes concurrent seeds converge.
		onReady: () => {
			if (chat.channels().length === 0)
				chat.createChannel({ id: 'general', name: 'General', topic: 'Company-wide announcements' });
		},
	});
	chat = createChatStore(session.doc, user);
	const tabs = createTabStore(session.doc, user, (id) =>
		chat.channels().some((c) => c.id === id && !c.archived),
	);

	const calls = new Map<string, CallSession>();
	return {
		user,
		session,
		chat,
		tabs,
		config,
		call(channelId) {
			const existing = calls.get(channelId);
			if (existing && existing.status !== 'left') return existing;
			const roomId = `call-${options.workspaceId}-${channelId}`.slice(0, 128);
			const call = createCallSession({
				roomId,
				user,
				signaling: signaling(roomId),
				iceServers: config.iceServers,
				...(config.iceTransportPolicy ? { iceTransportPolicy: config.iceTransportPolicy } : {}),
				...(options.mediaDevices ? { mediaDevices: options.mediaDevices } : {}),
				...(options.createPeerConnection
					? { createPeerConnection: options.createPeerConnection }
					: {}),
				...(options.createStream ? { createStream: options.createStream } : {}),
			});
			call.on('status', (s) => {
				if (s === 'connected') session.updatePresence({ callRoom: roomId });
				if (s === 'left') session.updatePresence({ callRoom: '' });
			});
			calls.set(channelId, call);
			return call;
		},
		setActiveChannel: (channelId) => session.updatePresence({ channelId: channelId ?? '' }),
		setTyping: (channelId) => session.updatePresence({ typingIn: channelId ?? '' }),
		setAvailability: (availability) => session.updatePresence({ availability }),
		destroy() {
			for (const c of calls.values()) c.leave();
			chat.destroy();
			session.destroy();
		},
	};
}
