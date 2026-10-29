// The framework-neutral client every UI binds to. `createTeamsClient` owns a workspace (chat,
// presence, calls) and exposes one immutable `TeamsState` snapshot plus plain action functions, in
// the `getState` / `subscribe` shape that React (`useSyncExternalStore`), Vue, Solid, Svelte and
// Angular can all adopt in a few lines. Anything a screen shows is in the state; anything a button
// does is an action. DOM-free: storage, fetch and media devices are injectable. New code.
import * as Y from 'yjs';
import { fromBase64, toBase64 } from '../collab/codec';
import { Emitter } from '../collab/emitter';
import { createClientId } from '../collab/identity';
import type { ConnectionStatus } from '../collab/provider';
import type { CallParticipant, CallSession, MediaDevicesLike } from './call';
import type { Attachment, Channel, Message } from './model';
import type { ActivityItem } from './activity';
import { createActivityFeed } from './activity-feed';
import { type MentionTarget, mentionTargetsFor } from './mentions';
import { mentionDisplayText, promoteMentions } from './mention-text';
import type { NotificationSettings } from './notifications';
import { createFileActions } from './files';
import { type ChatActions, createChatActions } from './chat-actions';
import { type ChatView, type ContactView, chatViews, contactViews } from './chat-view';
import { type MeetingActions, createMeetingActions } from './meeting-actions';
import { createMeetingReminders } from './meeting-reminders';
import { type Meeting, meetingCallKey, meetingVisibleTo } from './meetings';
import { localZone } from './time-zone';
import { shouldNotify } from './notifications';
import { createMessageTransfers, type MessageTransfer } from './message-transfer';
import { checkFileAbort, withFileAbort, type FileOperationOptions } from './file-transfer';
import { chatRoomId } from './chats';
import { fileRequestHeaders, uploadTeamsServerFile } from './server-file-storage';
import type { ChannelTab, TabContent } from './tabs';
import { channelThreads, type MessageThread } from './threads';
import { tabConversationId } from './tab-conversation';
import { createDraftStore, type ChatDraft, type DraftContext, type SavedDraft } from './drafts';
import {
	createThreadFollows,
	type FollowedThread,
	type ThreadFollowSettings,
} from './followed-threads';
import type { StreamLike } from './peer';
import type { PinnedMessage } from './pins';
import {
	type PresenceController,
	type PresenceView,
	createPresenceController,
} from './presence-controller';
import type { SchemaStatus } from './schema';
import { defaultPresenceSetting, type StatusDuration } from './presence';
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
} from './view';
import {
	type Availability,
	type TeamsWorkspace,
	type WorkspaceOptions,
	createTeamsWorkspace,
} from './workspace';

export interface StorageLike {
	getItem: (key: string) => string | null;
	setItem: (key: string, value: string) => void;
}
export interface UploadableFile {
	name: string;
	/** Original display name when name is a unique remote storage key. */
	originalName?: string;
	size: number;
	type?: string;
}
export type FileUploader = (
	file: UploadableFile,
	context: { workspaceId: string; signal?: AbortSignal },
) => Promise<{ url: string }>;

export interface TeamsClientOptions extends Omit<WorkspaceOptions, 'doc'> {
	/** Where to keep a local snapshot of the shared document and read markers. Optional. */
	storage?: StorageLike;
	/** Where attachments go. Default: the configured server's `/files` endpoint. */
	uploadFile?: FileUploader;
	/** Minutes without activity before showing Away (0 never; default 5). A saved choice wins. */
	idleMinutes?: number;
	fetch?: typeof fetch;
	/** The clock, for reminders; tests pass their own. Default `Date.now`. */
	now?: () => number;
	/** How often to check for due meeting reminders, in ms. Default 15 seconds; 0 turns it off. */
	reminderCheckMs?: number;
}

export type CallPhase = 'prejoin' | 'joining' | 'connected';

export interface MeetingView extends Meeting {
	/** Someone is in this meeting's call right now. */
	live: boolean;
	/** You organised it. */
	mine: boolean;
}

export interface CallView {
	/** The channel, chat or meeting call key the call belongs to. */
	channelId: string;
	channelName: string;
	/** What the title shows: a channel (# name), a private chat or a scheduled meeting. */
	kind: 'channel' | 'chat' | 'meeting';
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
	/** Whether this app may edit the workspace; `unsupported-schema` means a newer app wrote it. */
	schema: SchemaStatus;
	user: { id: string; name: string };
	/** The signed session grants moderation (hide others' messages, manage others' channels and tabs). */
	moderator: boolean;
	/** What others see for you: your chosen status, or Away when idle. */
	availability: Availability;
	/** Your chosen status and message with their expiry, idleness and the idle setting. */
	presence: PresenceView;
	channels: ChannelView[];
	/** Your 1:1 and group chats, most recent first (the Chat app). */
	chats: ChatView[];
	/** Everybody on the workspace roster except you, online first (New chat, attendees). */
	contacts: ContactView[];
	/** Meetings in your calendar (organiser, attendee or channel meeting), not cancelled. */
	meetings: MeetingView[];
	/** The zone the calendar renders in: this device's. */
	viewerZone: string;
	/** The selected channel or chat. */
	selectedChannelId: string;
	channel: Channel | null;
	messages: Message[];
	posts: Message[];
	/** Pinned messages of the selected channel, most recently pinned first. */
	pinned: PinnedMessage[];
	replyCounts: Record<string, number>;
	thread: MessageThread | null;
	threadFollowed: boolean;
	followedThreads: FollowedThread[];
	/** Mentions of you, replies, reactions and pins, newest first, with personal read state. */
	activity: ActivityItem[];
	activityUnread: number;
	notificationSettings: NotificationSettings;
	threadFollowSettings: ThreadFollowSettings;
	draft: ChatDraft;
	drafts: SavedDraft[];
	messageTransfers: MessageTransfer[];
	files: FileEntry[];
	/** Every file in every channel, newest first. */
	allFiles: FileEntry[];
	/** Shared tabs of the selected channel; selection remains local to each client. */
	tabs: ChannelTab[];
	canUploadFiles: boolean;
	people: PersonView[];
	/** Who the composer's @ picker offers in the selected channel (people, then @channel and @team). */
	mentionTargets: MentionTarget[];
	typing: string[];
	/** The message the composer is replying to or editing, if any. */
	replyingTo: Message | null;
	editing: Message | null;
	searchQuery: string;
	searchResults: SearchHit[];
	call: CallView | null;
}

export interface TeamsEvents {
	notice: string;
	notification: ActivityItem;
}

export interface TeamsClient extends ChatActions, MeetingActions {
	getState: () => TeamsState;
	/** Called after every state change (coalesced per tick); returns the unsubscribe function. */
	subscribe: (listener: () => void) => () => void;
	/**
	 * `notice`: a message to show the person. `notification`: a new activity item that settings and
	 * presence (not Do not disturb) allow to be announced; the host shows the toast, desktop
	 * notification and sound.
	 */
	on: Emitter<TeamsEvents>['on'];
	readonly workspace: TeamsWorkspace;
	select: (channelId: string) => void;
	createChannel: (name: string, topic?: string) => void;
	/** Rename a channel you created (or any channel as a moderator). */
	renameChannel: (id: string, name: string) => boolean;
	/** Archive a channel you created (or any channel as a moderator). */
	archiveChannel: (id: string) => boolean;
	addTab: (
		name: string,
		content: TabContent,
		options?: { postToChannel?: boolean },
	) => ChannelTab | null;
	/** Open or start a tab discussion in the selected channel. Preserves other thread drafts. */
	openTabConversation: (tabId: string) => boolean;
	renameTab: (id: string, name: string) => boolean;
	removeTab: (id: string) => boolean;
	/** Upload a uniquely named copy and post it in the specified channel. Never overwrites the source. */
	saveFileCopy: (
		channelId: string,
		file: UploadableFile & Blob,
		options?: FileOperationOptions,
	) => Promise<Attachment>;
	/** Upload files directly into the captured channel under unique storage names. */
	uploadFiles: (
		channelId: string,
		files: (UploadableFile & Blob)[],
		options?: FileOperationOptions,
	) => Promise<Attachment[]>;
	/** Create a blank native Excel workbook and share it in the captured channel. */
	createWorkbook: (
		channelId: string,
		name: string,
		options?: FileOperationOptions,
	) => Promise<Attachment>;
	/**
	 * Post (or, while editing, save) the composer text. `mentions` are the people picked in the
	 * picker; any other `@Name` that matches a channel member is a mention too.
	 */
	send: (input: {
		text: string;
		files?: (UploadableFile & Blob)[];
		mentions?: readonly MentionTarget[];
	}) => Promise<void>;
	setDraft: (draft: ChatDraft) => void;
	openDraft: (context: DraftContext) => boolean;
	discardDraft: (context: DraftContext) => void;
	cancelSend: (transferId: string) => void;
	startReply: (messageId: string) => void;
	openThread: (messageId: string) => void;
	closeThread: () => void;
	followThread: (channelId: string, messageId: string, followed: boolean) => boolean;
	/** Personal read state for a followed thread, independent of channel read markers. */
	markThreadRead: (channelId: string, messageId: string, read?: boolean) => boolean;
	setThreadFollowSettings: (settings: Partial<ThreadFollowSettings>) => void;
	/** Mark an activity item read (or unread). Personal, kept on this device. */
	markActivityRead: (id: string, read?: boolean) => boolean;
	markAllActivityRead: () => boolean;
	/** What to announce: all activity, mentions only or nothing, and the sound. Kept on this device. */
	setNotificationSettings: (settings: Partial<NotificationSettings>) => void;
	startEdit: (messageId: string) => void;
	cancelCompose: () => void;
	deleteMessage: (messageId: string) => void;
	/** Pin a message of the selected channel (any member may). Shared through the Yjs document. */
	pinMessage: (messageId: string) => boolean;
	unpinMessage: (messageId: string) => boolean;
	toggleReaction: (messageId: string, emoji: string) => void;
	notifyTyping: () => void;
	/** Choose a status; `duration` resets it to Available later (default never). */
	setAvailability: (value: Availability, duration?: StatusDuration) => void;
	/** Set (or with an empty string clear) the custom status message; `duration` expires it. */
	setStatusMessage: (message: string, duration?: StatusDuration) => void;
	/** Minutes without activity before Away (0 never). Remembered in storage. */
	setIdleMinutes: (minutes: number) => void;
	/** Report user input so idleness resets. Call it from pointer and key events (throttle freely). */
	noteActivity: () => void;
	search: (query: string) => void;
	/**
	 * The URL to open for an attachment. On your own server it asks for a short-lived link signed
	 * for that one file (`POST /files/link/<id>/<name>` with the token in the Authorization header),
	 * so the long-lived token never appears in a URL. Other URLs come back unchanged.
	 */
	fileUrl: (attachment: { url?: string | undefined }) => Promise<string | undefined>;
	openCall: (channelId?: string) => Promise<void>;
	/** Open the pre-join screen of a scheduled meeting's call; a meeting can be joined early. */
	joinMeeting: (meetingId: string) => Promise<boolean>;
	setPrejoin: (next: Partial<{ audio: boolean; video: boolean }>) => Promise<void>;
	joinCall: () => Promise<void>;
	leaveCall: () => void;
	toggleMic: () => void;
	toggleCamera: () => Promise<void>;
	toggleScreenShare: () => Promise<void>;
	toggleHand: () => void;
	/** Persist pending local snapshots before navigation, without closing the client. */
	flushStorage: () => void;
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
	const now = options.now ?? Date.now;
	// Each private chat has its own document and local snapshot.
	const chatDocs = new Map<string, Y.Doc>();
	let onChatDocUpdate = (): void => undefined;
	const chatDocKey = (id: string): string => STORAGE_DOC + workspaceId + ':chat:' + id;
	const ws = createTeamsWorkspace({
		...options,
		doc,
		createChatDoc: (chat) => {
			const chatDoc = restoreDoc(storage, chatDocKey(chat.id).slice(STORAGE_DOC.length));
			chatDocs.set(chat.id, chatDoc);
			chatDoc.on('update', () => onChatDocUpdate());
			return chatDoc;
		},
	});
	const follows = createThreadFollows(
		storage,
		`teams:followed:${encodeURIComponent(workspaceId)}:${encodeURIComponent(ws.user.id)}`,
	);
	const reminders = createMeetingReminders(
		storage,
		`teams:meeting-reminders:${encodeURIComponent(workspaceId)}:${encodeURIComponent(ws.user.id)}`,
	);
	const feed = createActivityFeed(
		storage,
		`teams:activity:${encodeURIComponent(workspaceId)}:${encodeURIComponent(ws.user.id)}`,
	);
	const notices = new Emitter<TeamsEvents>();
	const drafts = createDraftStore(
		storage,
		`teams:drafts:${encodeURIComponent(workspaceId)}:${encodeURIComponent(ws.user.id)}`,
	);
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

	const selectionKey = `teams:selected:${encodeURIComponent(workspaceId)}:${encodeURIComponent(ws.user.id)}`;
	let selected = '';
	try {
		selected = storage?.getItem(selectionKey)?.slice(0, 64) ?? '';
	} catch {
		/* Storage is optional. */
	}
	let query = '';
	let replyId = '';
	let editId = '';
	let presence: PresenceController | undefined;
	const presenceView = (): PresenceView =>
		presence?.view() ?? {
			...defaultPresenceSetting(),
			effective: 'available',
			idle: false,
			idleMinutes: 0,
		};
	let typingTimer: ReturnType<typeof setTimeout> | undefined;
	let saveTimer: ReturnType<typeof setTimeout> | undefined;
	let destroyed = false;

	interface ActiveCall {
		channelId: string;
		kind: 'channel' | 'chat' | 'meeting';
		/** Title override for a meeting call. */
		label?: string;
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
	let threadId = '';
	let activeDraftId = '';
	let transfers: ReturnType<typeof createMessageTransfers> | undefined;
	const draftContext = (): DraftContext => {
		const root =
			threadId ||
			(replyId ? channelThreads(messagesOf(selected)).thread(replyId)?.root.id : undefined);
		return {
			channelId: selected,
			...(activeDraftId ? { draftId: activeDraftId } : {}),
			...(root ? { threadId: root } : {}),
			...(editId ? { editId } : {}),
			...(replyId ? { replyTo: replyId } : {}),
		};
	};
	const peers = (): PeerLike[] => ws.session.peers() as unknown as PeerLike[];
	/** Channels and private chats: everything with messages. */
	const visibleChannels = (): Channel[] => ws.chat.channels().filter((c) => !c.archived);

	const compute = (): TeamsState => {
		const everything = visibleChannels();
		const channels = everything.filter((c) => c.kind !== 'direct');
		if (!everything.some((c) => c.id === selected)) {
			activeDraftId = threadId = replyId = editId = '';
			selected = channels[0]?.id ?? '';
			if (selected) ws.setActiveChannel(selected);
		}
		if (selected) lastRead[selected] = Date.now();
		const p = peers();
		const messages = selected ? messagesOf(selected) : [];
		const threads = channelThreads(messages);
		const thread = threads.thread(threadId);
		const channel = everything.find((c) => c.id === selected) ?? null;
		const activity = feed.update(
			{
				selfId: ws.user.id,
				channels: everything,
				messagesOf,
				reactions: ws.chat.reactionEvents(),
				pinsOf: ws.chat.pinned,
				isFollowed: (channelId, rootId) => follows.has(channelId, rootId),
			},
			presenceView().effective,
		);
		const people = peopleViews(
			{
				id: ws.user.id,
				name: ws.user.name,
				color: ws.session.identity.userColor,
				availability: presenceView().effective,
				statusMessage: presenceView().message,
				inCall: call?.phase === 'connected',
			},
			p,
		);
		const session = call?.session ?? null;
		const participants = session?.participants() ?? [];
		const channelName = call?.label ?? everything.find((c) => c.id === call?.channelId)?.name ?? '';
		const meetingList: MeetingView[] = ws.meetings
			.list()
			.filter((m) => meetingVisibleTo(m, ws.user.id))
			.map((m) => ({
				...m,
				mine: m.organizerId === ws.user.id,
				live:
					call?.channelId === meetingCallKey(m.id) ||
					p.some((peer) => peer.callRoom === callRoomId(workspaceId, meetingCallKey(m.id))),
			}));
		const chatCall =
			call?.session && call.phase === 'connected'
				? { inCallRoom: callRoomId(workspaceId, call.channelId) }
				: {};
		return {
			status: ws.session.status,
			mode: ws.config.mode,
			schema: ws.schema.status(),
			user: ws.user,
			moderator: ws.user.moderator === true,
			availability: presenceView().effective,
			presence: presenceView(),
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
			chats: chatViews({
				chats: ws.chat.chats(),
				messagesOf,
				selfId: ws.user.id,
				selfStatus: presenceView().effective,
				lastRead,
				peers: p,
				workspaceId,
				selectedId: selected,
				...chatCall,
			}),
			contacts: contactViews(ws.roster.list(), ws.user.id, p),
			meetings: meetingList,
			viewerZone: localZone(),
			selectedChannelId: selected,
			channel,
			messages,
			posts: threads.posts,
			pinned: selected ? ws.chat.pinned(selected) : [],
			replyCounts: threads.replyCounts,
			thread,
			threadFollowed:
				!!thread && follows.has(selected, thread.root.id, (id) => threads.thread(id)?.root.id),
			followedThreads: follows.view(everything, messagesOf, ws.user.id),
			activity: activity.items,
			activityUnread: activity.unread,
			notificationSettings: feed.settings.get(),
			threadFollowSettings: follows.settings(),
			messageTransfers: transfers?.list() ?? [],
			draft: drafts.read(
				draftContext(),
				mentionDisplayText(messages.find((message) => message.id === editId) ?? { text: '' }),
			),
			drafts: drafts
				.list()
				.filter((draft) => everything.some((channel) => channel.id === draft.context.channelId)),
			files: filesOf(messages),
			allFiles: allFilesOf(everything, messagesOf),
			tabs: selected ? ws.tabs.tabs(selected) : [],
			canUploadFiles: Boolean(
				options.uploadFile || (ws.config.mode === 'server' && ws.config.syncUrl),
			),
			people,
			mentionTargets: ws.chat.chat(selected)
				? mentionTargetsFor([], ws.chat.chat(selected)!.members, ws.user.id).filter(
						(t) => t.kind === 'user',
					)
				: mentionTargetsFor(
						messages,
						people.flatMap((p) => (p.userId ? [{ id: p.userId, name: p.name }] : [])),
						ws.user.id,
					),
			typing: selected ? typingNames(p, selected) : [],
			replyingTo: messages.find((m) => m.id === replyId) ?? null,
			editing: messages.find((m) => m.id === editId) ?? null,
			searchQuery: query,
			searchResults: searchMessages(everything, messagesOf, query),
			call: call
				? {
						channelId: call.channelId,
						channelName,
						kind: call.kind,
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
			for (const item of feed.drain()) notices.emit('notification', item);
		});
	};

	const saveSnapshot = (): void => {
		if (!storage || ws.schema.status().state === 'unsupported-schema') return;
		try {
			const text = toBase64(Y.encodeStateAsUpdate(doc));
			if (text.length <= MAX_SNAPSHOT_CHARS) storage.setItem(STORAGE_DOC + workspaceId, text);
			storage.setItem(STORAGE_READ + workspaceId, JSON.stringify(lastRead));
			storage.setItem(selectionKey, selected);
			const mine = new Set(ws.chat.chats().map((c) => c.id));
			for (const [id, chatDoc] of chatDocs) {
				// A chat you left is dropped from this device; the rest keep a local copy.
				const text = mine.has(id) ? toBase64(Y.encodeStateAsUpdate(chatDoc)) : '';
				if (text.length <= MAX_SNAPSHOT_CHARS) storage.setItem(chatDocKey(id), text);
			}
		} catch {
			// Blocked or over quota: the convenience is lost, the app keeps working.
		}
	};
	const scheduleSave = (): void => {
		if (!storage || destroyed) return;
		clearTimeout(saveTimer);
		saveTimer = setTimeout(saveSnapshot, 800);
	};
	doc.on('update', scheduleSave);
	onChatDocUpdate = scheduleSave;

	presence = createPresenceController({
		storage,
		storageKey: `teams:presence:${encodeURIComponent(workspaceId)}:${encodeURIComponent(ws.user.id)}`,
		idleMinutes: options.idleMinutes,
		publish: (status, message) => {
			ws.setAvailability(status);
			ws.setStatusMessage(message);
		},
		changed: refresh,
	});
	state = compute();

	const unsubs = [
		ws.chat.observe(refresh),
		ws.session.on('peers', refresh),
		ws.session.on('status', refresh),
		ws.schema.subscribe((s) => {
			if (s.state === 'unsupported-schema')
				notice(
					'This workspace was upgraded by a newer version of the app. It is read-only until you update.',
				);
			refresh();
		}),
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

	const upload = async (
		file: UploadableFile & Blob,
		signal?: AbortSignal,
		channelId?: string,
	): Promise<Attachment> => {
		checkFileAbort(signal);
		// A private chat's files are stored under its room so the server limits them to its members.
		const scopeId =
			channelId !== undefined && ws.chat.isChat(channelId)
				? chatRoomId(workspaceId, channelId)
				: workspaceId;
		const attachment: Attachment = {
			name: file.name,
			kind: 'other',
			size: file.size,
			...(file.type ? { mime: file.type } : {}),
		};
		try {
			if (options.uploadFile) {
				attachment.url = (
					await withFileAbort(
						() => options.uploadFile!(file, { workspaceId, ...(signal ? { signal } : {}) }),
						signal,
					)
				).url;
				return attachment;
			}
			attachment.url = (
				await uploadTeamsServerFile(ws.config, scopeId, file, {
					userId: ws.user.id,
					...(signal ? { signal } : {}),
					...(options.fetch ? { fetch: options.fetch } : {}),
				})
			).url;
		} catch (error) {
			checkFileAbort(signal);
			throw error;
		}
		return attachment;
	};
	transfers = createMessageTransfers({
		available: (channelId) =>
			!destroyed && visibleChannels().some((channel) => channel.id === channelId),
		canUpload: () => Boolean(options.uploadFile || server()),
		upload,
		changed: refresh,
	});

	/** The creator or a moderator may rename or archive a shared channel (not a private chat). */
	const canManageChannel = (id: string): boolean => {
		const channel = ws.chat.channels().find((c) => c.id === id);
		return (
			channel !== undefined &&
			channel.kind === 'channel' &&
			(channel.createdBy === ws.user.id || ws.user.moderator === true)
		);
	};
	const act =
		<A extends unknown[]>(fn: (...args: A) => void) =>
		(...args: A): void => {
			fn(...args);
			refresh();
		};
	const find = (id: string): Message | undefined => messagesOf(selected).find((m) => m.id === id);
	const markThreadRead = (channelId: string, messageId: string, read = true): boolean => {
		const index = channelThreads(messagesOf(channelId));
		const thread = index.thread(messageId);
		if (!thread) return false;
		let readAt = 0;
		if (read)
			for (const message of [thread.root, ...thread.replies]) readAt = Math.max(readAt, message.ts);
		const changed = follows.mark(
			channelId,
			thread.root.id,
			readAt,
			!read,
			(id) => index.thread(id)?.root.id,
		);
		if (changed) refresh();
		return changed;
	};
	const followAuthored = (message: Message | null): void => {
		if (!message || !follows.settings()[message.replyTo ? 'replied' : 'started']) return;
		const index = channelThreads(messagesOf(message.channelId));
		const thread = index.thread(message.id);
		if (
			thread &&
			follows.set(message.channelId, thread.root.id, true, (id) => index.thread(id)?.root.id)
		)
			markThreadRead(message.channelId, thread.root.id);
	};

	const selectId = (id: string): void => {
		if (id === selected) return;
		selected = id;
		activeDraftId = threadId = replyId = editId = '';
		ws.setActiveChannel(id);
		scheduleSave();
	};
	const chatActions = createChatActions({
		ws,
		select: (id) => (selectId(id), refresh()),
		notice,
		refresh,
	});
	const meetingActions: MeetingActions = createMeetingActions({
		ws,
		meetings: ws.meetings,
		notice,
		refresh,
		now,
	});
	const checkReminders = (): void => {
		if (destroyed) return;
		const settings = feed.settings.get();
		const status = presenceView().effective;
		const due = reminders.due(ws.meetings.list(), ws.user.id, now(), (item) =>
			shouldNotify(item, settings, status),
		);
		for (const item of due) notices.emit('notification', item);
	};
	const reminderMs = options.reminderCheckMs ?? 15_000;
	const reminderTimer = reminderMs > 0 ? setInterval(checkReminders, reminderMs) : undefined;
	(reminderTimer as { unref?: () => void } | undefined)?.unref?.();
	checkReminders();

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
		...chatActions,
		...meetingActions,
		addTab(name, content, options) {
			if (destroyed) return null;
			const tab = ws.tabs.add(selected, name, content);
			if (tab && options?.postToChannel) followAuthored(ws.chat.ensureTabConversation(tab));
			refresh();
			return tab;
		},
		openTabConversation(id) {
			if (destroyed) return false;
			const tab = ws.tabs.tabs(selected).find((entry) => entry.id === id);
			if (!tab) return false;
			const existing = messagesOf(selected).some((message) => message.id === tabConversationId(id));
			const message = ws.chat.ensureTabConversation(tab);
			if (!message) return false;
			if (!existing) followAuthored(message);
			activeDraftId = '';
			threadId = message.id;
			replyId = editId = '';
			markThreadRead(selected, message.id);
			refresh();
			return true;
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
		...createFileActions({
			available: (id) => !destroyed && visibleChannels().some((c) => c.id === id),
			canUpload: () => Boolean(options.uploadFile || server()),
			upload,
			post: (channelId, text, attachments) => {
				const message = ws.chat.post(channelId, { text, attachments });
				followAuthored(message);
				refresh();
				return Boolean(message);
			},
		}),
		select: act(selectId),
		renameChannel(id, name) {
			const ok = canManageChannel(id) && ws.chat.renameChannel(id, name);
			if (ok) refresh();
			else notice('You cannot rename this channel');
			return ok;
		},
		archiveChannel(id) {
			const ok = canManageChannel(id);
			if (ok) {
				ws.chat.archiveChannel(id, true);
				refresh();
			} else notice('You cannot archive this channel');
			return ok;
		},
		createChannel: act((name, topic = '') => {
			const channel = ws.chat.createChannel({ name, topic });
			if (channel) {
				activeDraftId = threadId = replyId = editId = '';
				selected = channel.id;
				ws.setActiveChannel(channel.id);
			}
		}),
		async send({ text, files = [], mentions = [] }) {
			if (!selected) return;
			files = [...files];
			if (!editId && !text.trim() && !files.length) return;
			const context = draftContext();
			const previous = drafts.read(context);
			if (previous.missingFiles.length) {
				notice('Reattach or discard the missing draft attachments before sending');
				return;
			}
			const sentDraft: ChatDraft = { text, files, missingFiles: [] };
			const revision = drafts.clear(context);
			const channelId = selected;
			const replyingTo = replyId;
			// The composer shows names; the stored text carries the "@" for older clients.
			const picks = [...mentions, ...(state.editing?.mentions ?? [])];
			const stored = promoteMentions(text, picks);
			clearTimeout(typingTimer);
			ws.setTyping(undefined);
			if (editId) {
				if (!ws.chat.edit(selected, editId, stored, [...picks, ...state.mentionTargets])) {
					drafts.restore(context, sentDraft, revision);
					notice('The edit was not applied; the draft was retained');
					return refresh();
				}
				editId = '';
				return refresh();
			}
			try {
				const publish = (attachments: Attachment[]) => {
					if (destroyed || !visibleChannels().some((channel) => channel.id === channelId))
						throw new Error('The channel is no longer available');
					const message = ws.chat.post(channelId, {
						text: stored,
						mentions: [...picks, ...state.mentionTargets],
						...(replyingTo ? { replyTo: replyingTo } : {}),
						attachments,
					});
					if (!message) throw new Error('The message could not be shared');
					followAuthored(message);
					if (selected === channelId && replyId === replyingTo) replyId = '';
				};
				if (files.length) await transfers!.upload(context, files, publish);
				else publish([]);
			} catch (error) {
				if (!drafts.restore(context, sentDraft, revision))
					drafts.set({ ...context, draftId: createClientId() }, sentDraft);
				notice(
					error instanceof Error && error.name === 'AbortError'
						? 'Message send canceled. The message was saved in Drafts.'
						: `Message was not sent: ${error instanceof Error ? error.message : 'Upload failed'}. Saved in Drafts.`,
				);
			} finally {
				refresh();
			}
		},
		cancelSend: (id) => transfers?.cancel(id),
		startReply: act((id) => {
			if (find(id)) [replyId, editId] = [id, ''];
		}),
		setDraft: act((draft) => {
			drafts.set(draftContext(), draft);
		}),
		discardDraft: act((context) => {
			drafts.clear(context);
		}),
		openDraft(context) {
			if (!visibleChannels().some((channel) => channel.id === context.channelId)) return false;
			const index = channelThreads(messagesOf(context.channelId));
			if (context.threadId && !index.thread(context.threadId)) return false;
			if (
				context.editId &&
				!messagesOf(context.channelId).some(
					(message) =>
						message.id === context.editId && message.authorId === ws.user.id && !message.deleted,
				)
			)
				return false;
			selected = context.channelId;
			activeDraftId = context.draftId ?? '';
			threadId = context.threadId ?? '';
			editId = context.editId ?? '';
			replyId = context.replyTo ?? '';
			ws.setActiveChannel(selected);
			refresh();
			return true;
		},
		openThread: act((id) => {
			const thread = channelThreads(messagesOf(selected)).thread(id);
			if (thread) {
				activeDraftId = '';
				threadId = thread.root.id;
				replyId = editId = '';
				markThreadRead(selected, thread.root.id);
			}
		}),
		closeThread: act(() => {
			activeDraftId = threadId = replyId = editId = '';
		}),
		followThread(channelId, messageId, followed) {
			const index = channelThreads(messagesOf(channelId));
			const thread = index.thread(messageId);
			if (followed && (!thread || !visibleChannels().some((channel) => channel.id === channelId)))
				return false;
			const changed = follows.set(
				channelId,
				thread?.root.id ?? messageId,
				followed,
				(id) => index.thread(id)?.root.id,
			);
			if (changed && followed && thread) markThreadRead(channelId, thread.root.id);
			refresh();
			return changed;
		},
		markThreadRead,
		markActivityRead(id, read = true) {
			const changed = feed.markRead(id, read);
			if (changed) refresh();
			return changed;
		},
		markAllActivityRead() {
			const changed = feed.markAllRead();
			if (changed) refresh();
			return changed;
		},
		setNotificationSettings: act((settings) => {
			feed.settings.set(settings);
		}),
		setThreadFollowSettings: act((settings) => {
			follows.configure(settings);
		}),
		startEdit: act((id) => {
			const m = find(id);
			if (m && m.authorId === ws.user.id) {
				activeDraftId = '';
				[editId, replyId] = [id, ''];
			}
		}),
		cancelCompose: act(() => {
			activeDraftId = replyId = editId = '';
		}),
		deleteMessage: act((id) => void ws.chat.remove(selected, id)),
		pinMessage(id) {
			const ok = Boolean(selected) && ws.chat.pin(selected, id);
			if (ok) refresh();
			else notice('This message cannot be pinned');
			return ok;
		},
		unpinMessage(id) {
			const ok = Boolean(selected) && ws.chat.unpin(selected, id);
			if (ok) refresh();
			return ok;
		},
		toggleReaction: act((id, emoji) => {
			const mine = find(id)?.reactions[emoji]?.includes(ws.user.id);
			if (mine) ws.chat.unreact(id, emoji, selected);
			else ws.chat.react(id, emoji, selected);
		}),
		notifyTyping() {
			if (!selected) return;
			ws.setTyping(selected);
			clearTimeout(typingTimer);
			typingTimer = setTimeout(() => ws.setTyping(undefined), TYPING_MS);
		},
		setAvailability: (value, duration) => presence?.setStatus(value, duration),
		setStatusMessage: (message, duration) => presence?.setMessage(message, duration),
		setIdleMinutes: (minutes) => presence?.setIdleMinutes(minutes),
		noteActivity: () => presence?.noteActivity(),
		search: act((q) => {
			query = q;
		}),
		async fileUrl(attachment) {
			const target = server();
			const raw = attachment.url;
			if (!raw || !target) return raw;
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
						headers: fileRequestHeaders(ws.config, ws.user.id),
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
				kind: ws.chat.isChat(channelId) ? 'chat' : 'channel',
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
		async joinMeeting(meetingId) {
			const meeting = ws.meetings.get(meetingId);
			if (!meeting || meeting.cancelled || !meetingVisibleTo(meeting, ws.user.id)) {
				notice('This meeting is not available');
				return false;
			}
			const key = meetingCallKey(meeting.id);
			if (call?.channelId === key) return true;
			leave();
			// The call's chat panel shows the meeting chat (or the channel of a channel meeting).
			const conversation = meeting.chatId ?? meeting.channelId;
			if (conversation && visibleChannels().some((c) => c.id === conversation))
				selectId(conversation);
			call = {
				channelId: key,
				kind: 'meeting',
				label: meeting.title,
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
			return true;
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
		flushStorage() {
			if (destroyed) return;
			clearTimeout(saveTimer);
			saveSnapshot();
		},
		destroy() {
			if (destroyed) return;
			saveSnapshot();
			destroyed = true;
			transfers?.destroy();
			clearTimeout(typingTimer);
			clearTimeout(saveTimer);
			clearInterval(reminderTimer);
			presence?.destroy();
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
