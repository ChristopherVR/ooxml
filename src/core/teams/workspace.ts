// The workspace ties the pieces together for one user: a collab session (chat state plus presence
// over a sync transport chosen from the server config), the chat store, and call sessions with
// signaling chosen the same way. This is what a viewer binding instantiates; everything above it is
// UI. DOM-free; browser classes are injectable. New code.
import type * as Y from 'yjs';
import { createBroadcastTransport } from '../collab/broadcast-transport';
import type { BroadcastChannelConstructor } from '../collab/broadcast-transport';
import { createClientId } from '../collab/identity';
import type { Transport } from '../collab/provider';
import { type CollabSession, createCollabSession } from '../collab/session';
import { transportProvider } from '../collab/transport-provider';
import { isValidId } from '../collab/validation';
import { createWebSocketTransport } from '../collab/websocket-transport';
import type { WebSocketConstructor } from '../collab/websocket-transport';
import { type CallSession, type MediaDevicesLike, createCallSession } from './call';
import { type ChatStore, createChatStore } from './chat';
import { type ChatRooms, createChatRooms } from './chat-rooms';
import { type ChatInfo, chatRoomId } from './chats';
import { type MeetingStore, createMeetingStore } from './meetings';
import { type Roster, createRoster } from './roster';
import { createTabStore, type TabStore } from './tabs';
import {
	type LegacyAvailability,
	type PresenceStatus,
	legacyAvailability,
	sanitizeStatusMessage,
	statusFromWire,
} from './presence';
import type { PeerConnectionFactory, StreamFactory } from './peer';
import { type TeamsServerConfig, endpointUrl, localServerConfig } from './server-config';
import {
	type SignalingChannel,
	createBroadcastSignaling,
	createWebSocketSignaling,
} from './signaling';

/** What a person shows as. Widened from the original three values; see `./presence`. */
export type Availability = PresenceStatus;

/** What each person publishes with their presence. */
export interface TeamsPresence {
	channelId?: string;
	/** Channel the person is typing in right now. */
	typingIn?: string;
	callRoom?: string;
	/** The person's user id, so others can @mention them before they have written anything. */
	userId?: string;
	/** The closest of the original three values, kept so older clients still read presence. */
	availability?: LegacyAvailability;
	/** The full status (Do not disturb, Be right back, Appear offline, ...). */
	status?: PresenceStatus;
	/** Custom status message. */
	statusMessage?: string;
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
	/** A document for a private chat, for example restored from a local snapshot. */
	createChatDoc?: (chat: ChatInfo) => Y.Doc;
}

export interface TeamsWorkspace {
	readonly user: { id: string; name: string };
	readonly session: CollabSession<TeamsPresence>;
	/** Channels and private chats behind one store; chats sync in rooms of their own. */
	readonly chat: ChatRooms;
	readonly roster: Roster;
	/** Scheduled meetings, shared through the workspace document. */
	readonly meetings: MeetingStore;
	readonly tabs: TabStore;
	readonly config: TeamsServerConfig;
	/** A call for a channel; call `join()` on it. One call per channel per workspace. */
	call: (channelId: string) => CallSession;
	setActiveChannel: (channelId: string | undefined) => void;
	setTyping: (channelId: string | undefined) => void;
	setAvailability: (value: Availability) => void;
	setStatusMessage: (message: string) => void;
	destroy: () => void;
}

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
	if (isValidId(raw.userId)) out.userId = raw.userId;
	out.status = statusFromWire(raw);
	out.availability = legacyAvailability(out.status);
	const statusMessage = sanitizeStatusMessage(raw.statusMessage);
	if (statusMessage) out.statusMessage = statusMessage;
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

	let channels!: ChatStore;
	const session = createCollabSession<TeamsPresence>({
		roomId: options.workspaceId,
		...(options.doc ? { doc: options.doc } : {}),
		provider: transportProvider({ transport: syncTransport(options.workspaceId) }),
		user: { name: user.name, ...(options.user.color ? { color: options.user.color } : {}) },
		initialPresence: { availability: 'available', status: 'available', userId: user.id },
		sanitizePayload: sanitizePresence,
		teardown: true,
		// Any writer may seed the default channel; the fixed id makes concurrent seeds converge.
		onReady: () => {
			if (channels.channels().length === 0)
				channels.createChannel({
					id: 'general',
					name: 'General',
					topic: 'Company-wide announcements',
				});
		},
	});
	channels = createChatStore(session.doc, user);
	const tabs = createTabStore(session.doc, user, (id) =>
		channels.channels().some((c) => c.id === id && !c.archived),
	);
	const roster = createRoster(session.doc, user);
	roster.register();
	const meetings = createMeetingStore(session.doc, user);
	// A private chat syncs in a room of its own. The server admits only the members listed in the
	// chat document, so the user id and the first member list ride along as query parameters.
	const chat = createChatRooms({
		doc: session.doc,
		base: channels,
		user,
		...(options.createChatDoc ? { createDoc: options.createChatDoc } : {}),
		openRoom: (info, doc) => {
			const roomId = chatRoomId(options.workspaceId, info.id);
			const members = info.members.map((m) => m.id).join(',');
			const transport: Transport =
				config.mode === 'server' && config.syncUrl
					? createWebSocketTransport({
							url: endpointUrl(config.syncUrl, roomId, config.token, { user: user.id, members }),
							...(wsCtor ? { WebSocket: wsCtor } : {}),
						})
					: createBroadcastTransport({ roomId, ...(bcCtor ? { BroadcastChannel: bcCtor } : {}) });
			const chatSession = createCollabSession({
				roomId,
				doc,
				provider: transportProvider({ transport }),
				user: { name: user.name },
			});
			return { destroy: () => chatSession.destroy() };
		},
	});

	const calls = new Map<string, CallSession>();
	return {
		user,
		session,
		chat,
		roster,
		meetings,
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
		setAvailability: (status) =>
			session.updatePresence({ status, availability: legacyAvailability(status) }),
		setStatusMessage: (message) =>
			session.updatePresence({ statusMessage: sanitizeStatusMessage(message) }),
		destroy() {
			for (const c of calls.values()) c.leave();
			chat.destroy();
			session.destroy();
		},
	};
}
