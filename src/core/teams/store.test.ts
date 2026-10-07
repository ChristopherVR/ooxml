import { createTeamsClient, type TeamsClient } from './store.js';
import { searchMessages, channelViews, filesOf, filterFiles } from './view.js';
import type { Channel, Message } from './model.js';

const tick = (ms = 0): Promise<void> => new Promise((r) => setTimeout(r, ms));
const clients: TeamsClient[] = [];
const make = (name: string, workspaceId: string, extra = {}): TeamsClient => {
	const c = createTeamsClient({ workspaceId, user: { id: name, name }, ...extra });
	clients.push(c);
	return c;
};
afterEach(() => {
	for (const c of clients.splice(0)) c.destroy();
});

describe('teams client', () => {
	it('flushes recent posts and shared tabs before a client is replaced', async () => {
		const values = new Map<string, string>();
		const storage = {
			getItem: (key: string) => values.get(key) ?? null,
			setItem: (key: string, value: string) => void values.set(key, value),
		};
		const client = make('ada', 'replace-persistence', { storage });
		client.createChannel('Project');
		await tick();
		await client.send({ text: 'Preserve this immediately' });
		client.addTab('Project site', { type: 'website', url: 'https://example.com' });
		client.destroy();
		client.destroy();
		const restored = make('ada', 'replace-persistence', { storage });
		await tick();
		const channel = restored.getState().channels.find((entry) => entry.name === 'Project')!;
		restored.select(channel.id);
		await tick();
		expect(restored.getState().messages.map((message) => message.text)).toEqual([
			'Preserve this immediately',
		]);
		expect(restored.getState().tabs[0]?.name).toBe('Project site');
	});
	it('retains an entire attachment message after a partial upload and retries with unique names', async () => {
		let offline = true;
		const names: string[] = [];
		const client = make('ada', 'chat-partial-upload', {
			uploadFile: async (file: { name: string }) => {
				names.push(file.name);
				if (offline && names.length === 2) throw new Error('Storage offline');
				return { url: `https://files.test/${file.name}` };
			},
		});
		client.createChannel('Project');
		await tick();
		const channel = client.getState().selectedChannelId;
		const root = client.workspace.chat.post(channel, { text: 'Original thread' })!;
		client.openThread(root.id);
		client.startReply(root.id);
		const files = ['Budget.xlsx', 'Notes.md'].map((name) =>
			Object.assign(new Blob([name]), { name }),
		);
		await client.send({ text: 'Both files', files });
		await tick();
		expect(client.getState().messages).toHaveLength(1);
		expect(client.getState().draft.text).toBe('Both files');
		expect(client.getState().draft.files).toEqual(files);
		expect(client.getState().messageTransfers).toEqual([]);
		offline = false;
		await client.send(client.getState().draft);
		await tick();
		expect(client.getState().messages[1]).toMatchObject({
			text: 'Both files',
			replyTo: root.id,
			attachments: [{ name: 'Budget.xlsx', kind: 'xlsx' }, { name: 'Notes.md' }],
		});
		expect(new Set(names).size).toBe(4);
		expect(client.getState().drafts).toEqual([]);
	});
	it('cancels ignored adapters and recovers the original send without overwriting newer text', async () => {
		let finish!: (result: { url: string }) => void;
		let signal: AbortSignal | undefined;
		const client = make('ada', 'chat-cancel-recovery', {
			uploadFile: (_file: unknown, context: { signal?: AbortSignal }) => {
				signal = context.signal;
				return new Promise<{ url: string }>((resolve) => {
					finish = resolve;
				});
			},
		});
		client.createChannel('Project');
		await tick();
		const file = Object.assign(new Blob(['xlsx']), { name: 'Budget.xlsx' });
		const sending = client.send({ text: 'Original send', files: [file] });
		await tick();
		expect(client.getState().messageTransfers[0]?.progress).toMatchObject({
			completed: 0,
			total: 1,
			fileName: 'Budget.xlsx',
		});
		client.setDraft({ text: 'Newer compose text', files: [], missingFiles: [] });
		client.cancelSend(client.getState().messageTransfers[0]!.id);
		await sending;
		await tick();
		expect(signal?.aborted).toBe(true);
		expect(client.getState().draft.text).toBe('Newer compose text');
		const recovered = client.getState().drafts.find((draft) => draft.text === 'Original send')!;
		expect(recovered.context.draftId).toBeTruthy();
		expect(recovered.files[0]).toBe(file);
		finish({ url: 'https://files.test/late.xlsx' });
		await tick();
		expect(client.getState().messages).toEqual([]);
		expect(client.openDraft(recovered.context)).toBe(true);
		await tick();
		expect(client.getState().draft.text).toBe('Original send');
		client.discardDraft(recovered.context);
		await tick();
		expect(client.getState().drafts.map((draft) => draft.text)).toEqual(['Newer compose text']);
	});
	it('rejects chat attachments without storage and never shares filename-only messages', async () => {
		const client = make('ada', 'chat-no-storage');
		client.createChannel('Project');
		await tick();
		const file = Object.assign(new Blob(['xlsx']), { name: 'Budget.xlsx' });
		await client.send({ text: 'Budget', files: [file] });
		await tick();
		expect(client.getState().messages).toEqual([]);
		expect(client.getState().draft.files[0]).toBe(file);
	});
	it('retains independent post, thread and edit drafts across navigation', async () => {
		const client = make('ada', 'draft-navigation');
		client.createChannel('Project');
		await tick();
		const channel = client.getState().selectedChannelId;
		const root = client.workspace.chat.post(channel, { text: 'Root' })!;
		const file = Object.assign(new Blob(['xlsx']), { name: 'Budget.xlsx' });
		client.setDraft({ text: 'New post draft', files: [file], missingFiles: [] });
		client.openThread(root.id);
		client.setDraft({ text: 'Thread draft', files: [], missingFiles: [] });
		client.startEdit(root.id);
		await tick();
		expect(client.getState().draft.text).toBe('Root');
		client.setDraft({ text: 'Edited draft', files: [], missingFiles: [] });
		client.cancelCompose();
		await tick();
		expect(client.getState().draft.text).toBe('Thread draft');
		client.closeThread();
		await tick();
		expect(client.getState().draft.text).toBe('New post draft');
		expect(client.getState().draft.files[0]).toBe(file);
		await client.send({ text: '' });
		await tick();
		expect(client.getState().draft.text).toBe('New post draft');
		client.createChannel('Other');
		client.setDraft({ text: 'Other channel', files: [], missingFiles: [] });
		client.select(channel);
		await tick();
		expect(client.getState().draft.text).toBe('New post draft');
		const edit = client.getState().drafts.find((draft) => draft.context.editId)!;
		expect(client.openDraft(edit.context)).toBe(true);
		await tick();
		expect(client.getState().draft.text).toBe('Edited draft');
	});
	it('rejects sending restored attachment names until they are reattached or discarded', async () => {
		const client = make('ada', 'draft-missing');
		client.createChannel('Project');
		await tick();
		client.setDraft({ text: 'Unsent', files: [], missingFiles: ['Budget.xlsx'] });
		await client.send({ text: 'Unsent' });
		await tick();
		expect(client.getState().messages).toEqual([]);
		expect(client.getState().draft.text).toBe('Unsent');
		client.setDraft({ text: 'Unsent', files: [], missingFiles: [] });
		await client.send({ text: 'Unsent' });
		await tick();
		expect(client.getState().messages[0]?.text).toBe('Unsent');
		expect(client.getState().drafts).toEqual([]);
	});
	it('keeps a newer draft while an earlier attachment send completes', async () => {
		let finish!: (result: { url: string }) => void;
		const client = make('ada', 'draft-pending-send', {
			uploadFile: () =>
				new Promise<{ url: string }>((resolve) => {
					finish = resolve;
				}),
		});
		client.createChannel('Project');
		await tick();
		const file = Object.assign(new Blob(['xlsx']), { name: 'Budget.xlsx' });
		client.setDraft({ text: 'First', files: [file], missingFiles: [] });
		const sending = client.send({ text: 'First', files: [file] });
		client.setDraft({ text: 'Second', files: [], missingFiles: [] });
		finish({ url: 'https://files.test/budget.xlsx' });
		await sending;
		await tick();
		expect(client.getState().draft.text).toBe('Second');
		expect(client.getState().messages[0]?.text).toBe('First');
	});
	it('automatically follows successful posts and replies according to personal settings', async () => {
		const client = make('ada', 'auto-follow');
		client.createChannel('Project');
		await tick();
		const channel = client.getState().selectedChannelId;
		await client.send({ text: 'Started thread' });
		await tick();
		const root = client.getState().posts[0]!;
		expect(client.getState().followedThreads[0]?.root.id).toBe(root.id);
		client.setThreadFollowSettings({ started: false });
		await client.send({ text: 'Unfollowed post' });
		await tick();
		expect(client.getState().followedThreads).toHaveLength(1);
		client.followThread(channel, root.id, false);
		client.startReply(root.id);
		await client.send({ text: 'Automatically followed reply' });
		await tick();
		expect(client.getState().followedThreads[0]?.root.id).toBe(root.id);
		client.setThreadFollowSettings({ replied: false });
		client.followThread(channel, root.id, false);
		client.startReply(root.id);
		await client.send({ text: 'Reply without following' });
		await tick();
		expect(client.getState().followedThreads).toEqual([]);
		client.setThreadFollowSettings({ started: true });
		await client.send({ text: '' });
		client.startEdit(root.id);
		await client.send({ text: 'Edited original post' });
		await tick();
		expect(client.getState().followedThreads).toEqual([]);
	});
	it('restores followed threads per user and workspace and follows the root', async () => {
		const values = new Map<string, string>();
		const storage = {
			getItem: (key: string) => values.get(key) ?? null,
			setItem: (key: string, value: string) => {
				values.set(key, value);
			},
		};
		const client = make('ada', 'follow-persistence', { storage });
		client.createChannel('Followed');
		await tick();
		const channel = client.getState().selectedChannelId;
		const root = client.workspace.chat.post(channel, { text: 'Root' })!;
		const reply = client.workspace.chat.post(channel, { text: 'Reply', replyTo: root.id })!;
		expect(client.followThread(channel, reply.id, true)).toBe(true);
		client.openThread(reply.id);
		await tick();
		expect(client.getState().threadFollowed).toBe(true);
		expect(client.getState().followedThreads[0]?.root.id).toBe(root.id);
		await tick(900);
		const restored = make('ada', 'follow-persistence', { storage });
		await tick();
		expect(restored.getState().followedThreads).toHaveLength(1);
		expect(restored.markThreadRead(channel, reply.id, false)).toBe(true);
		await tick();
		expect(restored.getState().followedThreads[0]?.unread).toBe(1);
		restored.openThread(reply.id);
		await tick();
		expect(restored.getState().followedThreads[0]?.unread).toBe(0);
		expect(make('bob', 'follow-persistence', { storage }).getState().followedThreads).toEqual([]);
		expect(make('ada', 'other-follow-workspace', { storage }).getState().followedThreads).toEqual(
			[],
		);
		expect(restored.followThread(channel, root.id, false)).toBe(true);
		await tick();
		expect(restored.getState().followedThreads).toEqual([]);
	});
	it('keeps thread selection local and resets it when switching channels', async () => {
		const client = make('ada', 'store-threads');
		client.createChannel('Source');
		await tick();
		const channel = client.getState().selectedChannelId;
		const root = client.workspace.chat.post(channel, { text: 'Root' })!;
		const reply = client.workspace.chat.post(channel, { text: 'Reply', replyTo: root.id })!;
		client.openThread(reply.id);
		await tick();
		expect(client.getState().thread?.root.id).toBe(root.id);
		expect(client.getState().posts.map((post) => post.id)).toEqual([root.id]);
		client.startReply(reply.id);
		await client.send({ text: 'Nested reply' });
		await tick();
		expect(client.getState().thread?.replies).toHaveLength(2);
		client.deleteMessage(root.id);
		await tick();
		expect(client.getState().thread?.root.deleted).toBe(true);
		client.createChannel('Other');
		await tick();
		expect(client.getState().thread).toBeNull();
		expect(client.getState().replyingTo).toBeNull();
	});
	it('passes cancellation to host storage and does not post a canceled copy', async () => {
		const controller = new AbortController();
		let seen: AbortSignal | undefined;
		let finish!: (result: { url: string }) => void;
		const client = make('ada', 'store-cancel-copy', {
			uploadFile: (_file: unknown, context: { signal?: AbortSignal }) => {
				seen = context.signal;
				return new Promise<{ url: string }>((resolve) => {
					finish = resolve;
				});
			},
		});
		client.createChannel('Files');
		await tick();
		const channel = client.getState().selectedChannelId;
		const pending = client.saveFileCopy(
			channel,
			Object.assign(new Blob(['x']), { name: 'Budget.xlsx' }),
			{ signal: controller.signal },
		);
		expect(seen).toBe(controller.signal);
		controller.abort();
		await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
		finish({ url: 'https://files.test/copy.xlsx' });
		await tick();
		expect(client.workspace.chat.messages(channel)).toEqual([]);
	});
	it('keeps a pending attachment message and reply in its original channel', async () => {
		let finish!: (result: { url: string }) => void;
		const a = make('ada', 'store-captured-send', {
			uploadFile: () => new Promise<{ url: string }>((resolve) => (finish = resolve)),
		});
		a.createChannel('Source');
		await tick();
		const source = a.getState().selectedChannelId;
		const message = a.workspace.chat.post(source, { text: 'Original discussion' })!;
		a.startReply(message.id);
		const pending = a.send({
			text: 'Attached budget',
			files: [Object.assign(new Blob(['xlsx']), { name: 'Budget.xlsx' })],
		});
		a.createChannel('Other');
		await tick();
		const other = a.getState().selectedChannelId;
		const reply = a.workspace.chat.post(other, { text: 'Other discussion' })!;
		a.startReply(reply.id);
		finish({ url: 'https://files.test/Budget.xlsx' });
		await pending;
		await tick();
		expect(a.workspace.chat.messages(source).at(-1)).toMatchObject({
			text: 'Attached budget',
			replyTo: message.id,
		});
		expect(a.workspace.chat.messages(other)).toHaveLength(1);
		expect(a.getState().replyingTo?.id).toBe(reply.id);
	});
	it('shares tabs while leaving each client selection local', async () => {
		const a = make('ada', 'store-tabs');
		const b = make('bob', 'store-tabs');
		a.createChannel('Project');
		await tick(200);
		const channel = a.getState().selectedChannelId;
		b.select(channel);
		const tab = a.addTab('Project site', { type: 'website', url: 'https://site.test' });
		await tick(200);
		expect(b.getState().tabs[0]?.id).toBe(tab?.id);
		expect(b.removeTab(tab!.id)).toBe(false);
		expect(a.renameTab(tab!.id, 'Design site')).toBe(true);
		await tick(100);
		expect(b.getState().tabs[0]?.name).toBe('Design site');
	});
	it('saves unique file copies to the captured channel without replacing the original', async () => {
		const uploaded: (Blob & { name: string })[] = [];
		const a = make('ada', 'store-copies', {
			uploadFile: async (file: Blob & { name: string }) => {
				uploaded.push(file);
				return { url: `https://files.test/${file.name}` };
			},
		});
		a.createChannel('Budget');
		await tick();
		const channelId = a.getState().selectedChannelId;
		const file = Object.assign(
			new Blob(['workbook'], {
				type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			}),
			{ name: 'Budget.xlsx' },
		);
		a.createChannel('Other');
		const first = await a.saveFileCopy(channelId, file);
		const second = await a.saveFileCopy(channelId, file);
		expect(first.name).toBe('Budget.xlsx');
		expect(first.kind).toBe('xlsx');
		expect(first.url).not.toBe(second.url);
		expect(
			uploaded
				.map((f) => f.name)
				.every((name) => name.startsWith('copy-') && name.endsWith('.xlsx')),
		).toBe(true);
		expect(await uploaded[0]!.text()).toBe('workbook');
		expect(a.workspace.chat.messages(channelId)).toHaveLength(2);
		expect(a.getState().messages).toHaveLength(0);
	});
	it('does not share a failed upload or a copy whose channel disappeared', async () => {
		const failed = make('ada', 'store-failed-copy', {
			uploadFile: async () => {
				throw new Error('Storage offline');
			},
		});
		failed.createChannel('Files');
		await tick();
		const file = Object.assign(new Blob(['workbook']), { name: 'Budget.xlsx' });
		await expect(failed.saveFileCopy(failed.getState().selectedChannelId, file)).rejects.toThrow(
			'could not be uploaded',
		);
		expect(failed.getState().files).toEqual([]);
		let resolve!: (value: { url: string }) => void;
		const pending = make('ada', 'store-removed-copy', {
			uploadFile: () =>
				new Promise<{ url: string }>((r) => {
					resolve = r;
				}),
		});
		pending.createChannel('Files');
		await tick();
		const channelId = pending.getState().selectedChannelId;
		const result = pending.saveFileCopy(channelId, file);
		pending.workspace.chat.archiveChannel(channelId);
		resolve({ url: 'https://files.test/copy.xlsx' });
		await expect(result).rejects.toThrow('not shared');
		expect(pending.workspace.chat.messages(channelId)).toEqual([]);
	});
	it('seeds General, posts, and notifies subscribers once per burst', async () => {
		const a = make('ada', 'store-seed');
		await tick(3300);
		expect(a.getState().channels.map((c) => c.name)).toEqual(['General']);
		expect(a.getState().selectedChannelId).toBe('general');
		let n = 0;
		a.subscribe(() => n++);
		await a.send({ text: 'hello' });
		await a.send({ text: 'again' });
		await tick();
		expect(a.getState().messages.map((m) => m.text)).toEqual(['hello', 'again']);
		expect(n).toBeGreaterThan(0);
	}, 10_000);

	it('supports reply, edit, delete and reaction toggling', async () => {
		const a = make('ada', 'store-actions');
		a.createChannel('Ops');
		await a.send({ text: 'one' });
		await tick();
		const id = a.getState().messages[0]!.id;
		a.startReply(id);
		await tick();
		expect(a.getState().replyingTo?.id).toBe(id);
		await a.send({ text: 'two' });
		await tick();
		expect(a.getState().messages[1]?.replyTo).toBe(id);
		a.startEdit(id);
		await tick();
		expect(a.getState().editing?.text).toBe('one');
		await a.send({ text: 'one (fixed)' });
		await tick();
		expect(a.getState().messages[0]).toMatchObject({ text: 'one (fixed)' });
		a.toggleReaction(id, '👍');
		await tick();
		expect(a.getState().messages[0]?.reactions['👍']).toEqual(['ada']);
		a.toggleReaction(id, '👍');
		await tick();
		expect(a.getState().messages[0]?.reactions['👍']).toBeUndefined();
		a.deleteMessage(id);
		await tick();
		expect(a.getState().messages[0]?.deleted).toBe(true);
	});

	it('shares state between two clients of one workspace and tracks unread', async () => {
		const a = make('ada', 'store-pair');
		const b = make('bob', 'store-pair');
		a.createChannel('Lobby');
		a.createChannel('Design', 'Pixels');
		await tick(200);
		const design = a.getState().channels.find((c) => c.name === 'Design')!;
		expect(b.getState().channels.some((c) => c.id === design.id)).toBe(true);
		await b.send({ text: 'from bob' });
		await tick(100);
		a.select(design.id);
		b.select(design.id);
		await a.send({ text: 'hi bob' });
		await tick(200);
		expect(b.getState().messages.map((m) => m.text)).toContain('hi bob');
		expect(b.getState().people.map((p) => p.name)).toContain('ada');
		b.select(b.getState().channels.find((c) => c.name === 'Lobby')!.id);
		await tick();
		await a.send({ text: 'ping' });
		await tick(200);
		expect(b.getState().channels.find((c) => c.id === design.id)?.unread).toBeGreaterThan(0);
	}, 10_000);

	it('searches across channels', async () => {
		const a = make('ada', 'store-search');
		a.createChannel('Ops');
		await a.send({ text: 'Deploy the NEW build' });
		a.search('new build');
		await tick();
		expect(a.getState().searchResults).toHaveLength(1);
		expect(a.getState().searchResults[0]?.channelName).toBe('Ops');
		a.search('');
		await tick();
		expect(a.getState().searchResults).toEqual([]);
	});

	it('walks the call flow: pre-join, join, leave', async () => {
		const track = (kind: string) => ({ kind, enabled: true, stop() {} });
		const stream = () => {
			const t: ReturnType<typeof track>[] = [];
			return { getTracks: () => t, addTrack: (x: ReturnType<typeof track>) => void t.push(x) };
		};
		const a = make('ada', 'store-call', {
			mediaDevices: {
				getUserMedia: async () => ({
					getTracks: () => [track('audio'), track('video')],
					addTrack() {},
				}),
			},
			createStream: stream,
			createPeerConnection: () => {
				throw new Error('no peers expected');
			},
		});
		a.createChannel('Standup');
		await a.openCall();
		expect(a.getState().call).toMatchObject({ phase: 'prejoin', channelName: 'Standup' });
		expect(a.getState().call?.preview).not.toBeNull();
		await a.setPrejoin({ video: false });
		expect(a.getState().call?.preview).toBeNull();
		await a.joinCall();
		await tick(50);
		expect(a.getState().call?.phase).toBe('connected');
		expect(a.getState().call?.self?.state.audio).toBe(true);
		expect(a.getState().channels.find((c) => c.name === 'Standup')?.live).toBe(true);
		a.leaveCall();
		await tick();
		expect(a.getState().call).toBeNull();
	});

	it('asks the server for a signed link and never puts the token in a URL', async () => {
		const calls: { url: string; auth: string | undefined }[] = [];
		const fetchStub = (async (url: string, init?: { headers?: Record<string, string> }) => {
			calls.push({ url, auth: init?.headers?.Authorization });
			return { ok: true, json: async () => ({ url: '/files/x/a.docx?sig=abc&exp=9' }) };
		}) as unknown as typeof fetch;
		const a = make('ada', 'store-token', {
			fetch: fetchStub,
			config: {
				mode: 'server',
				syncUrl: 'ws://127.0.0.1:1/sync',
				iceServers: [],
				token: 'sek ret',
			},
		});
		expect(await a.fileUrl({ url: 'http://127.0.0.1:1/files/x/a.docx' })).toBe(
			'http://127.0.0.1:1/files/x/a.docx?sig=abc&exp=9',
		);
		expect(calls).toEqual([
			{ url: 'http://127.0.0.1:1/files/link/x/a.docx', auth: 'Bearer sek ret' },
		]);
		// Foreign origins and other paths are returned as they are, and no request is made.
		expect(await a.fileUrl({ url: 'https://elsewhere.example/a.docx' })).toBe(
			'https://elsewhere.example/a.docx',
		);
		expect(await a.fileUrl({ url: 'http://127.0.0.1:1.evil.example/files/x/a.docx' })).toBe(
			'http://127.0.0.1:1.evil.example/files/x/a.docx',
		);
		expect(await a.fileUrl({})).toBeUndefined();
		expect(calls).toHaveLength(1);
	});
});

describe('view projections', () => {
	it('filters files by all search words without changing the source ordering', () => {
		const files = [
			{
				name: 'Budget.xlsx',
				kind: 'xlsx' as const,
				author: 'Ada',
				channelName: 'Finance',
				ts: 2,
				messageId: 'a',
			},
			{
				name: 'Notes.md',
				kind: 'other' as const,
				author: 'Bob',
				channelName: 'Finance',
				ts: 1,
				messageId: 'b',
			},
		];
		expect(filterFiles(files, 'BUDGET ada')).toEqual([files[0]]);
		expect(filterFiles(files, 'finance')).toEqual(files);
		expect(filterFiles(files, 'budget bob')).toEqual([]);
		expect(filterFiles(files, '  ')).toEqual(files);
	});
	const ch = (id: string, archived = false): Channel => ({
		id,
		name: id,
		topic: '',
		kind: 'channel',
		createdAt: 0,
		createdBy: 'x',
		archived,
	});
	const msg = (id: string, text: string, ts: number, authorId = 'x'): Message => ({
		id,
		channelId: 'c',
		authorId,
		authorName: authorId,
		text,
		ts,
		deleted: false,
		attachments: [],
		reactions: {},
	});

	it('counts unread from others after the read marker, never for the selected channel', () => {
		const messages = [msg('1', 'a', 5), msg('2', 'b', 20), msg('3', 'mine', 30, 'me')];
		const views = channelViews({
			channels: [ch('c'), ch('d'), ch('gone', true)],
			messagesOf: () => messages,
			selfId: 'me',
			lastRead: { c: 10 },
			peers: [],
			workspaceId: 'w',
			selectedId: 'd',
		});
		expect(views.map((v) => [v.id, v.unread])).toEqual([
			['c', 1],
			['d', 0],
		]);
	});
	it('lists files newest first and ignores deleted messages', () => {
		const a: Message = { ...msg('1', '', 1), attachments: [{ name: 'a.xlsx', kind: 'xlsx' }] };
		const b: Message = { ...msg('2', '', 2), attachments: [{ name: 'b.pptx', kind: 'pptx' }] };
		const gone: Message = {
			...msg('3', '', 3),
			deleted: true,
			attachments: [{ name: 'c.docx', kind: 'docx' }],
		};
		expect(filesOf([a, b, gone]).map((f) => f.name)).toEqual(['b.pptx', 'a.xlsx']);
	});
	it('finds by text or author, skipping archived channels', () => {
		const hits = searchMessages(
			[ch('c'), ch('old', true)],
			(id) => [msg(id, 'Budget review', 1, 'Zoe')],
			'zoe',
		);
		expect(hits.map((h) => h.channelId)).toEqual(['c']);
	});
});
