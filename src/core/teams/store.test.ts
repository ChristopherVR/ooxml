import { createTeamsClient, type TeamsClient } from './store.js';
import { searchMessages, channelViews, filesOf } from './view.js';
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
