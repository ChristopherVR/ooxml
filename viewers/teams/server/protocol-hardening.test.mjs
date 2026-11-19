// node --test server/protocol-hardening.test.mjs
// The server and protocol boundary against clients and peers that misbehave: weak session secrets,
// sessions that expire on open sockets, hook-only servers, call signaling, presence size, storage
// that new names would grow, and the cost of refusals and checks.
import assert from 'node:assert/strict';
import { existsSync, mkdtempSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, describe, it } from 'node:test';
import * as encoding from 'lib0/encoding';
import { createGroupChatId, shortHash } from 'ooxml-core/teams';
import { WebSocket } from 'ws';
import * as awarenessProtocol from 'y-protocols/awareness';
import * as Y from 'yjs';
import { resolveLimits } from './guards.mjs';
import { createTeamsServer } from './index.mjs';
import { storageName } from './room-store.mjs';
import { seedRoom, sleep, syncClient, until } from './test-support.mjs';

const USERS = { ann: 'pw-ann', bob: 'pw-bob', eve: 'pw-eve' };
const STRONG_SECRET = 'a3f9c27e61b84d05f1e2a9c3b7d60e4f8a1c5b9d2e7f3a6c';

const dirs = [];
const servers = [];
const sockets = [];
after(async () => {
	for (const ws of sockets) ws.terminate();
	for (const server of servers) await server.close();
	for (const dir of dirs) rmSync(dir, { recursive: true, force: true });
});

const tempDir = () => {
	const dir = mkdtempSync(join(tmpdir(), 'teams-f8-'));
	dirs.push(dir);
	return dir;
};

async function start(options = {}) {
	const server = createTeamsServer({
		dataDir: options.dataDir ?? tempDir(),
		users: USERS,
		sessionRate: { perIp: 1000, perUser: 1000 },
		log: () => {},
		...options,
	});
	servers.push(server);
	const { port } = await server.listen(0, '127.0.0.1');
	const http = `http://127.0.0.1:${port}`;
	const sessions = {};
	for (const user of Object.keys(options.users ?? USERS)) {
		const res = await fetch(`${http}/auth/session`, {
			method: 'POST',
			body: JSON.stringify({ user, credential: (options.users ?? USERS)[user] }),
		});
		if (res.status === 200) sessions[user] = (await res.json()).session;
	}
	return { server, port, http, sessions };
}

const track = (client) => (sockets.push(client.ws), client);

/** A signaling socket: `status` is `open` or the HTTP status, `got` the signals it received. */
function signal(ctx, room, session) {
	return new Promise((resolve) => {
		const protocols = ['openteams', ...(session ? [`openteams.session.${session}`] : [])];
		const ws = new WebSocket(`ws://127.0.0.1:${ctx.port}/signal/${room}`, protocols);
		sockets.push(ws);
		const peer = { ws, got: [], closed: null, status: 'connecting' };
		ws.on('message', (data) => peer.got.push(JSON.parse(String(data))));
		ws.on('close', (code, reason) => (peer.closed = { code, reason: String(reason) }));
		ws.on('open', () => resolve(Object.assign(peer, { status: 'open' })));
		ws.on('unexpected-response', (_req, res) =>
			resolve(Object.assign(peer, { status: res.statusCode })),
		);
		ws.on('error', () => {});
	});
}

const hello = (from) => JSON.stringify({ type: 'hello', from, name: from, state: {} });

/** A presence message carrying one state of about `bytes` bytes. */
function presence(bytes) {
	const awareness = new awarenessProtocol.Awareness(new Y.Doc());
	awareness.setLocalState({ note: 'x'.repeat(bytes) });
	const encoder = encoding.createEncoder();
	encoding.writeVarUint(encoder, 1);
	encoding.writeVarUint8Array(
		encoder,
		awarenessProtocol.encodeAwarenessUpdate(awareness, [awareness.clientID]),
	);
	awareness.destroy();
	return encoding.toUint8Array(encoder);
}

/** The longest the event loop went without running a timer while `work` ran. */
async function longestStall(work) {
	let last = performance.now();
	let longest = 0;
	const timer = setInterval(() => {
		const now = performance.now();
		longest = Math.max(longest, now - last);
		last = now;
	}, 2);
	try {
		await work();
		await sleep(20);
	} finally {
		clearInterval(timer);
	}
	return longest;
}

describe('session secrets', () => {
	it('refuses to start with a short or repetitive secret', () => {
		for (const secret of ['change-me', 'short', 'a'.repeat(64), '0123456'.repeat(5)])
			assert.throws(
				() => createTeamsServer({ dataDir: tempDir(), users: USERS, sessionSecret: secret }),
				/session secret .* too weak/u,
				secret,
			);
	});

	it('starts with a long random secret, or with none', async () => {
		for (const sessionSecret of [STRONG_SECRET, '']) {
			const server = createTeamsServer({ dataDir: tempDir(), users: USERS, sessionSecret });
			await server.close();
		}
	});
});

describe('an expiring session', () => {
	it('closes an open socket when its session expires, before it can write again', async () => {
		const ctx = await start({ sessionTtlMs: 400 });
		const writer = track(syncClient(ctx.port, ctx.sessions.ann, 'acme'));
		await writer.ready;
		assert.equal(writer.status, 'open');
		assert.ok(await until(() => writer.closed !== null, 3000), 'the socket closes at expiry');
		assert.deepEqual(writer.closed, { code: 1008, reason: 'session expired' });
	});
});

describe('a server with only an authorizeSync hook', () => {
	it('offers no sessions, files or call signaling, and keeps sync behind the hook', async () => {
		const server = createTeamsServer({
			dataDir: tempDir(),
			authorizeSync: () => 'write',
			log: () => {},
		});
		servers.push(server);
		const { port } = await server.listen(0, '127.0.0.1');
		const http = `http://127.0.0.1:${port}`;
		const session = await fetch(`${http}/auth/session`, {
			method: 'POST',
			body: JSON.stringify({ user: 'anyone' }),
		});
		assert.equal(session.status, 404);
		assert.equal(
			(await fetch(`${http}/files/ws1/a.txt`, { method: 'POST', body: 'x' })).status,
			404,
		);
		assert.equal((await fetch(`${http}/files/ws1/a.txt`)).status, 404);
		assert.equal((await fetch(`${http}/files/ws1`)).status, 404);
		assert.equal((await signal({ port }, 'call-ws1-general')).status, 404);
		const sync = track(syncClient(port, null, 'ws1'));
		await sync.ready;
		assert.equal(sync.status, 'open');
	});

	it('still opens those routes in an explicit dev setup', async () => {
		const server = createTeamsServer({
			dataDir: tempDir(),
			authorizeSync: () => 'write',
			insecureDev: true,
			log: () => {},
		});
		servers.push(server);
		const { port } = await server.listen(0, '127.0.0.1');
		assert.equal((await signal({ port }, 'call-ws1-general')).status, 'open');
	});
});

describe('call signaling', () => {
	it('binds each socket to its signed-in user', async () => {
		const ctx = await start();
		const room = 'call-acme-general';
		const ann = await signal(ctx, room, ctx.sessions.ann);
		const bob = await signal(ctx, room, ctx.sessions.bob);
		const eve = await signal(ctx, room, ctx.sessions.eve);
		eve.ws.send(hello('bob'));
		assert.ok(await until(() => eve.closed !== null), 'a socket speaking for someone else closes');
		assert.deepEqual(eve.closed, { code: 1008, reason: 'peer id is not the signed-in user' });
		const eve2 = await signal(ctx, room, ctx.sessions.eve);
		eve2.ws.send(hello('eve'));
		ann.ws.send(hello('ann'));
		ann.ws.send(
			JSON.stringify({
				type: 'candidate',
				from: 'ann',
				to: 'bob',
				candidate: { candidate: 'candidate:1 1 udp 1 192.0.2.1 5000 typ host' },
			}),
		);
		assert.ok(await until(() => bob.got.some((s) => s.type === 'candidate')));
		await sleep(50);
		assert.deepEqual(
			ann.got.map((s) => `${s.type}:${s.from}`).filter((s) => s.startsWith('hello')),
			['hello:eve'],
			'only real hellos reach the others',
		);
		assert.equal(eve2.got.filter((s) => s.type === 'candidate').length, 0);
		assert.equal(eve2.closed, null);
	});

	it('refuses a second socket for a user already in the call', async () => {
		const ctx = await start();
		const first = await signal(ctx, 'call-acme-general', ctx.sessions.ann);
		const second = await signal(ctx, 'call-acme-general', ctx.sessions.ann);
		assert.ok(await until(() => second.closed !== null));
		assert.deepEqual(second.closed, { code: 1008, reason: 'peer id in use' });
		first.ws.send(hello('ann'));
		const bob = await signal(ctx, 'call-acme-general', ctx.sessions.bob);
		await sleep(50);
		assert.equal(first.closed, null);
		assert.equal(bob.got.filter((s) => s.type === 'bye').length, 0);
	});

	it('admits only the members of a private chat to its call, and drops one who leaves', async () => {
		const ctx = await start();
		const chatId = createGroupChatId('ann');
		const chatRoom = `chat-${shortHash('acme')}-${chatId}`;
		const ann = track(syncClient(ctx.port, ctx.sessions.ann, chatRoom));
		await ann.ready;
		assert.ok(await until(() => ann.doc.getMap('teams:chat-members').get('ann') === true));
		ann.doc.getMap('teams:chat-members').set('bob', true);
		const call = `call-${chatRoom}`;
		const bobSync = track(syncClient(ctx.port, ctx.sessions.bob, chatRoom));
		await bobSync.ready;
		assert.equal(bobSync.status, 'open');
		assert.equal((await signal(ctx, call, ctx.sessions.eve)).status, 403);
		const annCall = await signal(ctx, call, ctx.sessions.ann);
		const bobCall = await signal(ctx, call, ctx.sessions.bob);
		assert.equal(annCall.status, 'open');
		assert.equal(bobCall.status, 'open');
		// The form an older client used names the chat but no room to check: refused.
		assert.equal((await signal(ctx, `call-acme-${chatId}`, ctx.sessions.ann)).status, 403);
		assert.equal(
			(await signal(ctx, call.toUpperCase().replace('CALL', 'call'), ctx.sessions.ann)).status,
			403,
		);
		assert.ok(await until(() => bobSync.doc.getMap('teams:chat-members').get('bob') === true));
		bobSync.doc.getMap('teams:chat-members').delete('bob');
		assert.ok(
			await until(() => bobCall.closed !== null),
			'leaving the chat ends the call for them',
		);
		assert.deepEqual(bobCall.closed, { code: 1008, reason: 'not a member' });
		assert.equal(annCall.closed, null);
	});

	it('checks the claimed id against the chat in a dev setup', async () => {
		const server = createTeamsServer({ dataDir: tempDir(), insecureDev: true, log: () => {} });
		servers.push(server);
		const { port } = await server.listen(0, '127.0.0.1');
		const chatRoom = `chat-${shortHash('acme')}-${createGroupChatId('ann')}`;
		const ann = track(syncClient(port, null, chatRoom, { query: '?user=ann' }));
		await ann.ready;
		assert.ok(await until(() => ann.doc.getMap('teams:chat-members').get('ann') === true));
		const outsider = await signal({ port }, `call-${chatRoom}`);
		outsider.ws.send(hello('eve'));
		assert.ok(await until(() => outsider.closed !== null));
		assert.deepEqual(outsider.closed, { code: 1008, reason: 'not a member' });
		const member = await signal({ port }, `call-${chatRoom}`);
		member.ws.send(hello('ann'));
		await sleep(50);
		assert.equal(member.closed, null);
	});
});

describe('presence size', () => {
	it('closes a socket whose presence state is too large and never replays it', async () => {
		const ctx = await start({ maxAwarenessBytes: 8 * 1024 });
		const sender = track(syncClient(ctx.port, ctx.sessions.ann, 'acme'));
		await sender.ready;
		sender.ws.send(presence(20 * 1024));
		assert.ok(await until(() => sender.closed !== null));
		assert.deepEqual(sender.closed, { code: 1009, reason: 'presence too large' });
		const joiner = track(syncClient(ctx.port, ctx.sessions.bob, 'acme'));
		await joiner.ready;
		await sleep(100);
		assert.equal(joiner.messages.filter((m) => m[0] === 1).length, 0);
	});

	it('caps the presence a room holds for every joiner', async () => {
		const ctx = await start({ maxAwarenessBytes: 8 * 1024, maxRoomAwarenessBytes: 12 * 1024 });
		const first = track(syncClient(ctx.port, ctx.sessions.ann, 'acme'));
		const second = track(syncClient(ctx.port, ctx.sessions.bob, 'acme'));
		await first.ready;
		await second.ready;
		first.ws.send(presence(7 * 1024));
		await sleep(50);
		second.ws.send(presence(7 * 1024));
		assert.ok(await until(() => second.closed !== null));
		assert.deepEqual(second.closed, { code: 1009, reason: 'presence too large' });
		assert.equal(first.closed, null);
		const joiner = track(syncClient(ctx.port, ctx.sessions.eve, 'acme'));
		await joiner.ready;
		await sleep(100);
		const received = joiner.messages.filter((m) => m[0] === 1).reduce((n, m) => n + m.length, 0);
		assert.ok(received < 12 * 1024, `a joiner received ${received} bytes of presence`);
	});

	it('limits the presence bytes one socket sends per second', async () => {
		const ctx = await start({ maxAwarenessBytes: 8 * 1024, awarenessByteRate: 16 * 1024 });
		const sender = track(syncClient(ctx.port, ctx.sessions.ann, 'acme'));
		await sender.ready;
		for (let i = 0; i < 20 && sender.ws.readyState === 1; i++) sender.ws.send(presence(6 * 1024));
		assert.ok(await until(() => sender.closed !== null));
		assert.deepEqual(sender.closed, { code: 1008, reason: 'rate limit' });
	});
});

describe('storage that new names would grow', () => {
	it('has a default ceiling for files, rooms and each uploader', () => {
		const limits = resolveLimits({}, {});
		for (const key of ['filesQuotaBytes', 'roomsQuotaBytes', 'userQuotaBytes', 'maxStoredRooms'])
			assert.ok(Number.isFinite(limits[key]), key);
	});

	it('counts what one person uploads in every scope together', async () => {
		const ctx = await start({ userQuotaBytes: 1000 });
		for (const room of ['acme', 'other']) await seedRoom(ctx.port, ctx.sessions.ann, room);
		const post = (scope, user, body) =>
			fetch(`${ctx.http}/files/${scope}/f.bin`, {
				method: 'POST',
				headers: { 'X-Teams-Session': ctx.sessions[user] },
				body,
			});
		assert.equal((await post('acme', 'ann', 'x'.repeat(600))).status, 201);
		assert.equal((await post('other', 'ann', 'x'.repeat(600))).status, 507);
		assert.equal((await post('other', 'bob', 'x'.repeat(600))).status, 201);
	});

	it('takes uploads only for a workspace that exists when accounts are on', async () => {
		const ctx = await start();
		const post = (scope) =>
			fetch(`${ctx.http}/files/${scope}/f.bin`, {
				method: 'POST',
				headers: { 'X-Teams-Session': ctx.sessions.ann },
				body: 'x',
			});
		assert.equal((await post('nobody-made-this')).status, 404);
		await seedRoom(ctx.port, ctx.sessions.ann, 'acme');
		assert.equal((await post('acme')).status, 201);
	});

	it('stores no more rooms than allowed and nothing for a room nobody wrote to', async () => {
		const dataDir = tempDir();
		const ctx = await start({ dataDir, maxStoredRooms: 2 });
		assert.equal(await seedRoom(ctx.port, ctx.sessions.ann, 'one'), null);
		assert.equal(await seedRoom(ctx.port, ctx.sessions.ann, 'two'), null);
		assert.deepEqual(await seedRoom(ctx.port, ctx.sessions.ann, 'three'), {
			code: 1009,
			reason: 'server storage full',
		});
		const idle = track(syncClient(ctx.port, ctx.sessions.ann, 'idle'));
		await idle.ready;
		idle.close();
		await sleep(50);
		await ctx.server.close();
		const stored = readdirSync(join(dataDir, 'rooms')).filter((f) => f.endsWith('.bin'));
		assert.deepEqual(
			stored.sort(),
			[`${storageName('one')}.bin`, `${storageName('two')}.bin`].sort(),
		);
		assert.equal(existsSync(join(dataDir, 'rooms', `${storageName('idle')}.bin`)), false);
	});
});

/** A room of `count` messages by Ann, each about `chars` characters long. */
const messages = (count, chars) => (doc) => {
	const rows = doc.getMap('teams:messages:general');
	const text = 'lorem ipsum '.repeat(Math.ceil(chars / 12)).slice(0, chars);
	for (let i = 0; i < count; i++) {
		const row = new Y.Map();
		row.set('authorId', 'ann');
		row.set('text', `${i} ${text}`);
		row.set('ts', i);
		rows.set(`m${i}`, row);
	}
};

describe('the cost of checks and refusals', () => {
	it('checks an update in time that does not grow with the room', async () => {
		const ctx = await start({ messageRate: 1000, messageBurst: 1000 });
		assert.equal(await seedRoom(ctx.port, ctx.sessions.ann, 'acme', messages(2500, 2000)), null);
		const ann = track(syncClient(ctx.port, ctx.sessions.ann, 'acme'));
		await ann.ready;
		assert.ok(await until(() => ann.doc.getMap('teams:messages:general').size > 0, 10_000));
		const rows = ann.doc.getMap('teams:messages:general');
		const stall = await longestStall(async () => {
			for (let i = 0; i < 40; i++) {
				const row = new Y.Map();
				row.set('authorId', 'ann');
				row.set('text', `small ${i}`);
				rows.set(`n${i}`, row);
			}
			const bob = track(syncClient(ctx.port, ctx.sessions.bob, 'acme'));
			await bob.ready;
			assert.ok(await until(() => bob.doc.getMap('teams:messages:general').has('n39'), 10_000));
		});
		assert.equal(ann.closed, null);
		assert.ok(stall < 400, `the event loop stalled ${Math.round(stall)} ms`);
	});

	it('refuses a join to a room on disk without reading the room', async () => {
		const dataDir = tempDir();
		const first = await start({ dataDir });
		const chatRoom = `chat-${shortHash('acme')}-${createGroupChatId('ann')}`;
		assert.equal(
			await seedRoom(first.port, first.sessions.ann, chatRoom, messages(60_000, 40)),
			null,
		);
		await first.server.close();
		// A fresh process: the room is on disk only.
		const ctx = await start({ dataDir });
		const dial = async (user) => {
			const client = track(syncClient(ctx.port, ctx.sessions[user], chatRoom));
			await client.ready;
			return client.status;
		};
		const stall = await longestStall(async () => {
			assert.equal(await dial('eve'), 'http-403');
			assert.equal(await dial('eve'), 'http-403');
			const files = await fetch(`${ctx.http}/files/${chatRoom}`, {
				headers: { 'X-Teams-Session': ctx.sessions.eve },
			});
			assert.equal(files.status, 403);
			assert.equal(await dial('eve'), 'http-403');
		});
		assert.ok(stall < 60, `a refusal stalled the event loop ${Math.round(stall)} ms`);
		assert.equal(await dial('ann'), 'open');
	});

	it('turns away a client whose refusals each cost a room read', async () => {
		const dataDir = tempDir();
		const first = await start({ dataDir });
		const rooms = [1, 2, 3, 4].map(() => `chat-${shortHash('acme')}-${createGroupChatId('ann')}`);
		for (const room of rooms)
			assert.equal(await seedRoom(first.port, first.sessions.ann, room), null);
		await first.server.close();
		// Rooms as an older version stored them: no member list beside the snapshot.
		for (const entry of readdirSync(join(dataDir, 'rooms')))
			if (entry.endsWith('.members.json')) rmSync(join(dataDir, 'rooms', entry));
		const ctx = await start({ dataDir, maxRefusals: 3 });
		const dial = async (user, room) => {
			const client = track(syncClient(ctx.port, ctx.sessions[user], room));
			await client.ready;
			return client.status;
		};
		for (const room of rooms.slice(0, 3)) assert.equal(await dial('eve', room), 'http-403');
		assert.equal(await dial('eve', rooms[3]), 'http-429');
		assert.equal(await dial('eve', 'acme'), 'http-429');
		const files = await fetch(`${ctx.http}/files/${rooms[3]}`, {
			headers: { 'X-Teams-Session': ctx.sessions.eve },
		});
		assert.equal(files.status, 429);
		assert.equal(await dial('ann', rooms[3]), 'open', 'other people are not turned away');
	});

	it('turns away a client whose updates keep being refused', async () => {
		const ctx = await start({ maxRefusals: 2 });
		await seedRoom(ctx.port, ctx.sessions.ann, 'acme');
		for (let i = 0; i < 2; i++) {
			const eve = track(syncClient(ctx.port, ctx.sessions.eve, 'acme'));
			await eve.ready;
			const row = new Y.Map();
			row.set('authorId', 'ann');
			row.set('text', 'not mine');
			eve.doc.getMap('teams:messages:general').set(`x${i}`, row);
			assert.ok(await until(() => eve.closed !== null));
			assert.match(eve.closed.reason, /^update rejected/u);
		}
		const again = track(syncClient(ctx.port, ctx.sessions.eve, 'acme'));
		await again.ready;
		assert.equal(again.status, 'http-429');
	});
});
