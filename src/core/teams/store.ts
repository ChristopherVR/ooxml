// The framework-neutral client every UI binds to. `createTeamsClient` owns a workspace (chat,
// presence, calls) and exposes one immutable `TeamsState` snapshot plus plain action functions, in
// the `getState` / `subscribe` shape that React (`useSyncExternalStore`), Vue, Solid, Svelte and
// Angular can all adopt in a few lines. Anything a screen shows is in the state; anything a button
// does is an action. DOM-free: storage, fetch and media devices are injectable. New code.
import * as Y from 'yjs';
import { fromBase64, toBase64 } from '../collab/codec.js';
import { Emitter } from '../collab/emitter.js';
import type { ConnectionStatus } from '../collab/provider.js';
import type { CallParticipant, CallSession, MediaDevicesLike } from './call.js';
import type { Attachment, Channel, Message } from './model.js';
import { sanitizeAttachment, sanitizeChannelName } from './model.js';
import type { ChannelTab, TabContent } from './tabs.js';
import type { StreamLike } from './peer.js';
import {
	type ChannelView,
	type FileEntry,
	type PeerLike,
	type PersonView,
	type SearchHit,
	allFilesOf,
	channelViews,
	callRoomId,
	filesOf,
	peopleViews,
	searchMessages,
	typingNames,
} from './view.js';
import {
	type Availability,
	type TeamsWorkspace,
	type WorkspaceOptions,
	createTeamsWorkspace,
} from './workspace.js';

export interface StorageLike {
	getItem: (key: string) => string | null;
	setItem: (key: string, value: string) => void;
}
export interface UploadableFile {
	name: string;
	size: number;
	type?: string;
}
export type FileUploader = (
	file: UploadableFile,
	context: { workspaceId: string },
) => Promise<{ url: string }>;

export interface TeamsClientOptions extends Omit<WorkspaceOptions, 'doc'> {
	/** Where to keep a local snapshot of the shared document and read markers. Optional. */
	storage?: StorageLike;
	/** Where attachments go. Default: the configured server's `/files` endpoint. */
	uploadFile?: FileUploader;
	fetch?: typeof fetch;
}

export type CallPhase = 'prejoin' | 'joining' | 'connected';

export interface CallView {
	channelId: string;
	channelName: string;
	phase: CallPhase;
	/** What the pre-join screen will join with. */
	prejoin: { audio: boolean; video: boolean };
	/** Camera preview while on the pre-join screen. */
	preview: StreamLike | null;
	participants: CallParticipant[];
	self: CallParticipant | undefined;
	startedAt: number | null;
}

export interface TeamsState {
	status: ConnectionStatus;
	mode: 'local' | 'server';
	user: { id: string; name: string };
	availability: Availability;
	channels: ChannelView[];
	selectedChannelId: string;
	channel: Channel | null;
	messages: Message[];
	files: FileEntry[];
	/** Every file in every channel, newest first. */
	allFiles: FileEntry[];
	/** Shared tabs of the selected channel; selection remains local to each client. */
	tabs: ChannelTab[];
	canUploadFiles: boolean;
	people: PersonView[];
	typing: string[];
	/** The message the composer is replying to or editing, if any. */
	replyingTo: Message | null;
	editing: Message | null;
	searchQuery: string;
	searchResults: SearchHit[];
	call: CallView | null;
}

export interface TeamsClient {
	getState: () => TeamsState;
	/** Called after every state change (coalesced per tick); returns the unsubscribe function. */
	subscribe: (listener: () => void) => () => void;
	on: Emitter<{ notice: string }>['on'];
	readonly workspace: TeamsWorkspace;
	select: (channelId: string) => void;
	createChannel: (name: string, topic?: string) => void;
	addTab: (name: string, content: TabContent) => ChannelTab | null;
	renameTab: (id: string, name: string) => boolean;
	removeTab: (id: string) => boolean;
	/** Upload a uniquely named copy and post it in the specified channel. Never overwrites the source. */
	saveFileCopy: (channelId: string, file: UploadableFile & Blob) => Promise<Attachment>;
	send: (input: { text: string; files?: (UploadableFile & Blob)[] }) => Promise<void>;
	startReply: (messageId: string) => void;
	startEdit: (messageId: string) => void;
	cancelCompose: () => void;
	deleteMessage: (messageId: string) => void;
	toggleReaction: (messageId: string, emoji: string) => void;
	notifyTyping: () => void;
	setAvailability: (value: Availability) => void;
	search: (query: string) => void;
	/**
	 * The URL to open for an attachment. On your own server it asks for a short-lived link signed
	 * for that one file (`POST /files/link/<id>/<name>` with the token in the Authorization header),
	 * so the long-lived token never appears in a URL. Other URLs come back unchanged.
	 */
	fileUrl: (attachment: { url?: string | undefined }) => Promise<string | undefined>;
	openCall: (channelId?: string) => Promise<void>;
	setPrejoin: (next: Partial<{ audio: boolean; video: boolean }>) => Promise<void>;
	joinCall: () => Promise<void>;
	leaveCall: () => void;
	toggleMic: () => void;
	toggleCamera: () => Promise<void>;
	toggleScreenShare: () => Promise<void>;
	toggleHand: () => void;
	destroy: () => void;
}

const STORAGE_DOC = 'teams:doc:';
const STORAGE_READ = 'teams:read:';
const MAX_SNAPSHOT_CHARS = 3_000_000;
const TYPING_MS = 4000;

function restoreDoc(storage: StorageLike | undefined, id: string): Y.Doc {
	const doc = new Y.Doc();
	try {
		const bytes = fromBase64(storage?.getItem(STORAGE_DOC + id) ?? '');
		if (bytes) Y.applyUpdate(doc, bytes);
	} catch {
		// A corrupt snapshot is ignored; the room or peers re-supply the state.
	}
	return doc;
}

export function createTeamsClient(options: TeamsClientOptions): TeamsClient {
	const { storage, workspaceId } = options;
	const doc = restoreDoc(storage, workspaceId);
	const ws = createTeamsWorkspace({ ...options, doc });
	const notices = new Emitter<{ notice: string }>();
	const notice = (text: string): void => notices.emit('notice', text);
	const listeners = new Set<() => void>();
	let lastRead: Record<string, number> = {};
	try {
		lastRead = JSON.parse(storage?.getItem(STORAGE_READ + workspaceId) ?? '{}') as Record<
			string,
			number
		>;
	} catch {
		lastRead = {};
	}

	let selected = '';
	let query = '';
	let replyId = '';
	let editId = '';
	let availability: Availability = 'available';
	let typingTimer: ReturnType<typeof setTimeout> | undefined;
	let saveTimer: ReturnType<typeof setTimeout> | undefined;
	let destroyed = false;

	interface ActiveCall {
		channelId: string;
		phase: CallPhase;
		prejoin: { audio: boolean; video: boolean };
		preview: StreamLike | null;
		session: CallSession | null;
		startedAt: number | null;
		off: (() => void)[];
	}
	let call: ActiveCall | null = null;
	const media = (): MediaDevicesLike | undefined =>
		options.mediaDevices ??
		(globalThis as unknown as { navigator?: { mediaDevices?: MediaDevicesLike } }).navigator
			?.mediaDevices;

	const messagesOf = (id: string): Message[] => ws.chat.messages(id);
	const peers = (): PeerLike[] => ws.session.peers() as unknown as PeerLike[];
	const visibleChannels = (): Channel[] => ws.chat.channels().filter((c) => !c.archived);

	const compute = (): TeamsState => {
		const channels = visibleChannels();
		if (!channels.some((c) => c.id === selected)) {
			selected = channels[0]?.id ?? '';
			if (selected) ws.setActiveChannel(selected);
		}
		if (selected) lastRead[selected] = Date.now();
		const p = peers();
		const messages = selected ? messagesOf(selected) : [];
		const channel = channels.find((c) => c.id === selected) ?? null;
		const session = call?.session ?? null;
		const participants = session?.participants() ?? [];
		const channelName = channels.find((c) => c.id === call?.channelId)?.name ?? '';
		return {
			status: ws.session.status,
			mode: ws.config.mode,
			user: ws.user,
			availability,
			channels: channelViews({
				channels,
				messagesOf,
				selfId: ws.user.id,
				lastRead,
				peers: p,
				workspaceId,
				selectedId: selected,
				...(call?.session && call.phase === 'connected'
					? { inCallRoom: callRoomId(workspaceId, call.channelId) }
					: {}),
			}),
			selectedChannelId: selected,
			channel,
			messages,
			files: filesOf(messages),
			allFiles: allFilesOf(channels, messagesOf),
			tabs: selected ? ws.tabs.tabs(selected) : [],
			canUploadFiles: Boolean(
				options.uploadFile || (ws.config.mode === 'server' && ws.config.syncUrl),
			),
			people: peopleViews(
				{
					id: ws.user.id,
					name: ws.user.name,
					color: ws.session.identity.userColor,
					availability,
					inCall: call?.phase === 'connected',
				},
				p,
			),
			typing: selected ? typingNames(p, selected) : [],
			replyingTo: messages.find((m) => m.id === replyId) ?? null,
			editing: messages.find((m) => m.id === editId) ?? null,
			searchQuery: query,
			searchResults: searchMessages(channels, messagesOf, query),
			call: call
				? {
						channelId: call.channelId,
						channelName,
						phase: call.phase,
						prejoin: call.prejoin,
						preview: call.preview,
						participants,
						self: participants.find((x) => x.self),
						startedAt: call.startedAt,
					}
				: null,
		};
	};

	let state = compute();
	let queued = false;
	const refresh = (): void => {
		if (queued || destroyed) return;
		queued = true;
		queueMicrotask(() => {
			queued = false;
			if (destroyed) return;
			state = compute();
			for (const l of [...listeners]) l();
		});
	};

	const scheduleSave = (): void => {
		if (!storage) return;
		clearTimeout(saveTimer);
		saveTimer = setTimeout(() => {
			try {
				const text = toBase64(Y.encodeStateAsUpdate(doc));
				if (text.length <= MAX_SNAPSHOT_CHARS) storage.setItem(STORAGE_DOC + workspaceId, text);
				storage.setItem(STORAGE_READ + workspaceId, JSON.stringify(lastRead));
			} catch {
				// Blocked or over quota: the convenience is lost, the app keeps working.
			}
		}, 800);
	};
	doc.on('update', scheduleSave);

	const unsubs = [
		ws.chat.observe(refresh),
		ws.session.on('peers', refresh),
		ws.session.on('status', refresh),
		ws.session.on('error', (e) => notice(e.message)),
	];

	const stopPreview = (): void => {
		for (const t of call?.preview?.getTracks() ?? []) t.stop();
		if (call) call.preview = null;
	};
	const startPreview = async (): Promise<void> => {
		if (!call || call.phase !== 'prejoin' || call.preview || !call.prejoin.video) return;
		try {
			const stream = await media()?.getUserMedia({ video: true });
			if (!stream) return;
			if (call?.phase === 'prejoin' && call.prejoin.video) call.preview = stream;
			else for (const t of stream.getTracks()) t.stop();
		} catch (error) {
			notice(error instanceof Error ? error.message : 'Camera unavailable');
			if (call) call.prejoin = { ...call.prejoin, video: false };
		}
	};

	const server = (): { base: string; token: string | undefined } | null => {
		const { config } = ws;
		if (config.mode !== 'server' || !config.syncUrl) return null;
		const u = new URL(config.syncUrl);
		u.protocol = u.protocol === 'wss:' ? 'https:' : 'http:';
		return { base: u.origin, token: config.token };
	};

	const upload = async (file: UploadableFile & Blob): Promise<Attachment> => {
		const attachment: Attachment = {
			name: file.name,
			kind: 'other',
			size: file.size,
			...(file.type ? { mime: file.type } : {}),
		};
		try {
			if (options.uploadFile) {
				attachment.url = (await options.uploadFile(file, { workspaceId })).url;
				return attachment;
			}
			const target = server();
			if (!target) {
				notice('No file server configured: sharing the name only');
				return attachment;
			}
			const doFetch = options.fetch ?? globalThis.fetch;
			const res = await doFetch(
				`${target.base}/files/${workspaceId}/${encodeURIComponent(file.name)}`,
				{
					method: 'POST',
					body: file,
					...(target.token ? { headers: { Authorization: `Bearer ${target.token}` } } : {}),
				},
			);
			if (!res.ok) throw new Error(`Upload failed (${res.status})`);
			attachment.url = `${target.base}${((await res.json()) as { url: string }).url}`;
		} catch (error) {
			notice(error instanceof Error ? error.message : 'Upload failed');
		}
		return attachment;
	};

	const act =
		<A extends unknown[]>(fn: (...args: A) => void) =>
		(...args: A): void => {
			fn(...args);
			refresh();
		};
	const find = (id: string): Message | undefined => messagesOf(selected).find((m) => m.id === id);

	const leave = (): void => {
		const active = call;
		if (!active) return;
		stopPreview();
		active.session?.leave();
		for (const f of active.off) f();
		call = null;
	};

	return {
		getState: () => state,
		subscribe(listener) {
			listeners.add(listener);
			return () => listeners.delete(listener);
		},
		on: (event, listener) => notices.on(event, listener),
		workspace: ws,
		addTab(name, content) {
			const tab = ws.tabs.add(selected, name, content);
			refresh();
			return tab;
		},
		renameTab(id, name) {
			const changed = ws.tabs.rename(id, name);
			refresh();
			return changed;
		},
		removeTab(id) {
			const changed = ws.tabs.remove(id);
			refresh();
			return changed;
		},
		async saveFileCopy(channelId, file) {
			if (destroyed || !visibleChannels().some((c) => c.id === channelId))
				throw new Error('The channel is no longer available');
			if (!options.uploadFile && !server())
				throw new Error('Configure file storage before saving a copy');
			const name = sanitizeChannelName(file.name);
			if (!name || file.size > 33_554_432)
				throw new Error('This file cannot be saved to the channel');
			const ext = name.includes('.') ? `.${name.split('.').pop()}` : '';
			const unique = Object.assign(new Blob([file], { type: file.type ?? '' }), {
				name: `copy-${crypto.randomUUID()}${ext}`,
			});
			const uploaded = await upload(unique);
			const attachment = sanitizeAttachment({ ...uploaded, name });
			if (!attachment?.url) throw new Error('The copy could not be uploaded');
			if (destroyed || !visibleChannels().some((c) => c.id === channelId))
				throw new Error('The channel is no longer available; the uploaded copy was not shared');
			if (!ws.chat.post(channelId, { text: `Saved a copy of ${name}`, attachments: [attachment] }))
				throw new Error('The copy could not be shared');
			refresh();
			return attachment;
		},
		select: act((id) => {
			if (id === selected) return;
			selected = id;
			replyId = editId = '';
			ws.setActiveChannel(id);
			scheduleSave();
		}),
		createChannel: act((name, topic = '') => {
			const channel = ws.chat.createChannel({ name, topic });
			if (channel) {
				selected = channel.id;
				ws.setActiveChannel(channel.id);
			}
		}),
		async send({ text, files = [] }) {
			if (!selected) return;
			clearTimeout(typingTimer);
			ws.setTyping(undefined);
			if (editId) {
				if (!ws.chat.edit(selected, editId, text)) notice('Only your own messages can be edited');
				editId = '';
				return refresh();
			}
			const attachments: Attachment[] = [];
			for (const file of files) attachments.push(await upload(file));
			ws.chat.post(selected, { text, ...(replyId ? { replyTo: replyId } : {}), attachments });
			replyId = '';
			refresh();
		},
		startReply: act((id) => {
			if (find(id)) [replyId, editId] = [id, ''];
		}),
		startEdit: act((id) => {
			const m = find(id);
			if (m && m.authorId === ws.user.id) [editId, replyId] = [id, ''];
		}),
		cancelCompose: act(() => {
			replyId = editId = '';
		}),
		deleteMessage: act((id) => void ws.chat.remove(selected, id)),
		toggleReaction: act((id, emoji) => {
			const mine = find(id)?.reactions[emoji]?.includes(ws.user.id);
			if (mine) ws.chat.unreact(id, emoji);
			else ws.chat.react(id, emoji);
		}),
		notifyTyping() {
			if (!selected) return;
			ws.setTyping(selected);
			clearTimeout(typingTimer);
			typingTimer = setTimeout(() => ws.setTyping(undefined), TYPING_MS);
		},
		setAvailability: act((value) => {
			availability = value;
			ws.setAvailability(value);
		}),
		search: act((q) => {
			query = q;
		}),
		async fileUrl(attachment) {
			const target = server();
			const raw = attachment.url;
			if (!raw || !target?.token) return raw;
			// Attachment URLs come from peers: the token goes only to the exact configured origin
			// (a prefix test would match `http://host:8787.evil.example`), only for /files/ paths,
			// and never to a URL carrying credentials.
			let url: URL;
			try {
				url = new URL(raw);
			} catch {
				return raw;
			}
			if (url.origin !== target.base || url.username || url.password) return raw;
			if (!url.pathname.startsWith('/files/')) return raw;
			try {
				const doFetch = options.fetch ?? globalThis.fetch;
				const res = await doFetch(
					`${target.base}/files/link${url.pathname.slice('/files'.length)}`,
					{
						method: 'POST',
						headers: { Authorization: `Bearer ${target.token}` },
					},
				);
				if (!res.ok) throw new Error(`Link failed (${res.status})`);
				const { url: signed } = (await res.json()) as { url: string };
				return `${target.base}${signed}`;
			} catch (error) {
				notice(error instanceof Error ? error.message : 'Could not open the file');
				return undefined;
			}
		},
		async openCall(channelId = selected) {
			if (!channelId) return;
			if (call?.channelId === channelId) return;
			leave();
			if (channelId !== selected) {
				selected = channelId;
				ws.setActiveChannel(channelId);
			}
			call = {
				channelId,
				phase: 'prejoin',
				prejoin: { audio: true, video: true },
				preview: null,
				session: null,
				startedAt: null,
				off: [],
			};
			refresh();
			await startPreview();
			refresh();
		},
		async setPrejoin(next) {
			if (!call || call.phase !== 'prejoin') return;
			call.prejoin = { ...call.prejoin, ...next };
			if (next.video === false) stopPreview();
			refresh();
			await startPreview();
			refresh();
		},
		async joinCall() {
			const active = call;
			if (!active || active.phase !== 'prejoin') return;
			stopPreview();
			active.phase = 'joining';
			const session = ws.call(active.channelId);
			active.session = session;
			active.off.push(
				session.on('participants', refresh),
				session.on('error', (e) => notice(e.message)),
				session.on('status', (s) => {
					if (s === 'connected') {
						active.phase = 'connected';
						active.startedAt ??= Date.now();
					}
					if (s === 'left' && call === active) {
						for (const f of active.off) f();
						call = null;
					}
					refresh();
				}),
			);
			refresh();
			await session.join(active.prejoin);
			refresh();
		},
		leaveCall: act(leave),
		toggleMic: act(() => call?.session?.toggleMic()),
		async toggleCamera() {
			await call?.session?.toggleCamera();
			refresh();
		},
		async toggleScreenShare() {
			const s = call?.session;
			if (!s) return;
			const sharing = s.participants().find((p) => p.self)?.state.screen;
			await (sharing ? s.stopScreenShare() : s.startScreenShare());
			refresh();
		},
		toggleHand: act(() => call?.session?.toggleHand()),
		destroy() {
			destroyed = true;
			clearTimeout(typingTimer);
			clearTimeout(saveTimer);
			try {
				storage?.setItem(STORAGE_READ + workspaceId, JSON.stringify(lastRead));
			} catch {
				// storage unavailable
			}
			leave();
			for (const u of unsubs) u();
			doc.off('update', scheduleSave);
			listeners.clear();
			ws.destroy();
		},
	};
}
