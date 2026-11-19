// node --test server/chat-access.test.mjs
// Private chats: only the people listed in a chat's room may join it, and nobody but the people a
// chat id names can create the room.
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, before, describe, it } from 'node:test';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import {
	chatRoomId,
	createGroupChatId,
	createTeamsClient,
	directChatId,
	legacyDirectChatId,
	shortHash,
} from 'ooxml-core/teams';
import { WebSocket } from 'ws';
import * as syncProtocol from 'y-protocols/sync';
import * as Y from 'yjs';
import * as chatAccess from './chat-access.mjs';
import { createTeamsServer } from './index.mjs';

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const until = async (fn, ms = 4000) => {
	const end = Date.now() + ms;
	while (Date.now() < end) {
		if (await fn()) return true;
		await sleep(25);
	}
	return false;
};

describe('chat room admission rules', () => {
	const room = (chatId) => `chat-${shortHash('acme')}-${chatId}`;
	const none = new Set();
	const attempt = (name, user, members, existing = none, started = false) =>
		chatAccess.authorizeChatJoin({ name, user, members, existing, started });

	it('derives ids exactly like the client', () => {
		assert.equal(chatAccess.shortHash('ü€'), shortHash('ü€'));
		assert.equal(chatAccess.directChatId('bob', 'ada'), directChatId('ada', 'bob'));
	});

	it('lets only the creator named by a group id start its room, alone', () => {
		const name = room(createGroupChatId('ada'));
		assert.deepEqual(attempt(name, 'ada'), { ok: true, seed: ['ada'] });
		assert.deepEqual(attempt(name, 'ada', 'ada,bob,cy'), { ok: true, seed: ['ada'] });
		assert.equal(attempt(name, 'bob', 'ada,bob').ok, false);
		assert.equal(attempt(name, 'eve', 'eve,ada,bob').ok, false);
	});

	it('does not start a room from an id that names nobody', () => {
		for (const id of ['g-1', 'mt-meeting', legacyDirectChatId('ada', 'bob')])
			assert.equal(attempt(room(id), 'ada', 'ada,bob').ok, false, id);
		assert.equal(attempt('plain-room', 'ada').ok, false);
	});

	it('only creates a 1:1 room for the two people its id was derived from', () => {
		const name = room(directChatId('ada', 'bob'));
		assert.deepEqual(attempt(name, 'ada', 'ada,bob'), { ok: true, seed: ['ada', 'bob'] });
		assert.equal(attempt(name, 'bob', 'bob,ada').ok, true);
		assert.equal(attempt(name, 'eve', 'eve,ada,bob').ok, false, 'a third person cannot squat');
		assert.equal(attempt(name, 'eve', 'ada,bob').ok, false);
		assert.equal(attempt(name, 'eve', 'eve,bob').ok, false);
		assert.equal(attempt(name, 'ada', 'ada').ok, false);
	});

	it('admits members of a started room and never widens or restarts it', () => {
		const name = room(createGroupChatId('ada'));
		const known = new Set(['ada', 'bob']);
		assert.deepEqual(attempt(name, 'bob', '', known, true), { ok: true, seed: [] });
		assert.equal(attempt(name, 'eve', 'eve,ada,bob', known, true).ok, false);
		assert.equal(attempt(name, 'ada', '', none, true).ok, false, 'an emptied room stays closed');
	});
});

describe('private chats on the reference server', () => {
	let dir;
	let server;
	let port;
	const clients = [];
	const sockets = [];
	const config = () => ({
		mode: 'server',
		syncUrl: `ws://127.0.0.1:${port}/sync`,
		signalingUrl: `ws://127.0.0.1:${port}/signal`,
		iceServers: [],
	});
	const client = (name, workspaceId, moderator = false) => {
		const c = createTeamsClient({
			workspaceId,
			user: { id: name, name },
			moderator,
			config: config(),
			WebSocket,
			reminderCheckMs: 0,
		});
		clients.push(c);
		return c;
	};
	/** Resolves with 'open' or the HTTP status the server answered the upgrade with. */
	const dial = (workspaceId, chatId, query) =>
		new Promise((resolve) => {
			const ws = new WebSocket(
				`ws://127.0.0.1:${port}/sync/${chatRoomId(workspaceId, chatId)}?${new URLSearchParams(query)}`,
			);
			sockets.push(ws);
			const closed = new Promise((r) => ws.on('close', (code) => r(code)));
			ws.on('open', () => resolve({ status: 'open', ws, closed }));
			ws.on('unexpected-response', (_req, res) => resolve({ status: res.statusCode }));
			ws.on('error', () => resolve({ status: 0 }));
		});
	before(async () => {
		dir = mkdtempSync(join(tmpdir(), 'teams-chat-access-'));
		server = createTeamsServer({ dataDir: dir, insecureDev: true, maxConnectionsPerIp: 200 });
		port = (await server.listen(0, '127.0.0.1')).port;
	});
	after(async () => {
		for (const c of clients) c.destroy();
		for (const ws of sockets) ws.terminate();
		await server.close();
		rmSync(dir, { recursive: true, force: true });
	});

	it('syncs a 1:1 chat between its two people and keeps a third person out of the room', async () => {
		const workspace = 'chat-ws-1';
		const ada = client('ada', workspace);
		const bob = client('bob', workspace);
		const eve = client('eve', workspace);
		assert.ok(await until(() => ada.getState().contacts.length === 2));
		const id = ada.openDirectChat('bob');
		await ada.send({ text: 'for bob only' });
		assert.ok(await until(() => bob.getState().chats.some((c) => c.id === id)));
		bob.select(id);
		assert.ok(await until(() => bob.getState().messages.some((m) => m.text === 'for bob only')));
		await bob.send({ text: 'got it' });
		assert.ok(await until(() => ada.getState().messages.some((m) => m.text === 'got it')));

		// Eve sees the directory row in the workspace but her client never opens the chat.
		assert.deepEqual(eve.getState().chats, []);
		assert.deepEqual(eve.workspace.chat.messages(id), []);
		// Even a hand-made socket is refused: wrong user, a widened member list, no user at all.
		assert.equal((await dial(workspace, id, { user: 'eve' })).status, 403);
		assert.equal((await dial(workspace, id, { user: 'eve', members: 'eve,ada,bob' })).status, 403);
		assert.equal((await dial(workspace, id, {})).status, 403);
		const insider = await dial(workspace, id, { user: 'bob' });
		assert.equal(insider.status, 'open');
		insider.ws.close();
	});

	it('refuses a squatter on a new 1:1 room and persists membership across a restart', async () => {
		const workspace = 'chat-ws-2';
		const id = directChatId('ada', 'bob');
		assert.equal((await dial(workspace, id, { user: 'eve', members: 'eve,ada,bob' })).status, 403);
		assert.equal((await dial(workspace, id, { user: 'eve', members: 'eve,bob' })).status, 403);
		const first = await dial(workspace, id, { user: 'ada', members: 'ada,bob' });
		assert.equal(first.status, 'open');
		first.ws.close();
		await sleep(700);
		assert.equal((await dial(workspace, id, { user: 'eve', members: 'eve,ada,bob' })).status, 403);
		assert.equal((await dial(workspace, id, { user: 'bob' })).status, 'open');
	});

	it('runs a group chat from creation through add, leave and removal', async () => {
		const workspace = 'chat-ws-3';
		const ada = client('ada', workspace, true);
		const bob = client('bob', workspace);
		client('cy', workspace);
		const dan = client('dan', workspace);
		assert.ok(await until(() => ada.getState().contacts.length === 3));
		const id = ada.createGroupChat(['bob', 'cy'], 'Crew');
		await ada.send({ text: 'welcome' });
		assert.ok(await until(() => bob.getState().chats.some((c) => c.id === id)));
		bob.select(id);
		assert.ok(await until(() => bob.getState().messages.some((m) => m.text === 'welcome')));

		ada.addToChat(id, ['dan']);
		assert.ok(await until(() => dan.getState().chats.some((c) => c.id === id)));
		dan.select(id);
		assert.ok(await until(() => dan.getState().messages.some((m) => m.text === 'welcome')));

		bob.leaveChat(id);
		assert.ok(await until(async () => (await dial(workspace, id, { user: 'bob' })).status === 403));

		const cySocket = await dial(workspace, id, { user: 'cy' });
		assert.equal(cySocket.status, 'open');
		ada.removeFromChat(id, 'cy');
		assert.equal(await Promise.race([cySocket.closed, sleep(4000).then(() => 'still open')]), 1008);
		assert.equal((await dial(workspace, id, { user: 'cy' })).status, 403);
		assert.equal((await dial(workspace, id, { user: 'dan' })).status, 'open');
	});

	describe('leaving while offline', () => {
		/** A WebSocket class whose connections can be cut, and fail while they are. */
		function network() {
			const net = { offline: false, live: new Set() };
			net.WebSocket = class extends WebSocket {
				constructor(url, protocols) {
					super(net.offline ? 'ws://127.0.0.1:9/' : url, protocols);
					net.live.add(this);
					this.on('close', () => net.live.delete(this));
				}
			};
			net.cut = () => {
				net.offline = true;
				for (const ws of net.live) ws.close();
			};
			return net;
		}
		const memoryStorage = () => {
			const items = new Map();
			return {
				getItem: (key) => items.get(key) ?? null,
				setItem: (key, value) => void items.set(key, value),
			};
		};
		const device = (name, workspaceId, net, storage) => {
			const c = createTeamsClient({
				workspaceId,
				user: { id: name, name },
				config: config(),
				WebSocket: net.WebSocket,
				...(storage ? { storage } : {}),
				reminderCheckMs: 0,
			});
			clients.push(c);
			return c;
		};
		/** Ada's group with bob and cy; cy went offline and left longer ago than rooms linger. */
		async function leftOffline(workspace, storage) {
			const net = network();
			const ada = client('ada', workspace);
			client('bob', workspace);
			const cy = device('cy', workspace, net, storage);
			assert.ok(await until(() => ada.getState().contacts.length === 2));
			const id = ada.createGroupChat(['bob', 'cy'], 'Crew');
			await ada.send({ text: 'welcome' });
			assert.ok(await until(() => cy.getState().chats.some((c) => c.id === id)));
			await sleep(200);
			net.cut();
			await sleep(100);
			assert.equal(cy.leaveChat(id), true);
			await sleep(1800);
			const stillListed = await dial(workspace, id, { user: 'cy' });
			assert.equal(stillListed.status, 'open', 'the server has not heard yet');
			stillListed.ws.close();
			return { id, net, ada, cy };
		}
		const refused = (workspace, id) =>
			until(async () => {
				const attempt = await dial(workspace, id, { user: 'cy' });
				attempt.ws?.close();
				await sleep(100);
				return attempt.status === 403;
			}, 15_000);

		it('takes effect in the room once the client is back', async () => {
			const workspace = 'chat-ws-offline-1';
			const { id, net, ada } = await leftOffline(workspace);
			net.offline = false;
			assert.ok(await refused(workspace, id), 'the room no longer admits them');
			assert.ok(await until(() => ada.getState().chats[0]?.members.length === 2));
			ada.select(id);
			assert.ok(
				await until(() => ada.getState().messages.some((m) => m.text === 'cy left the chat')),
			);
		});

		it('survives closing the app before the client is back', async () => {
			const workspace = 'chat-ws-offline-2';
			const storage = memoryStorage();
			const { id, net, cy } = await leftOffline(workspace, storage);
			cy.destroy();
			net.offline = false;
			device('cy', workspace, net, storage);
			assert.ok(await refused(workspace, id), 'the room no longer admits them');
		});
	});

	it('gives a meeting its chat under an id only the organizer can start', async () => {
		const workspace = 'chat-ws-4';
		const ada = client('ada', workspace);
		const bob = client('bob', workspace);
		assert.ok(await until(() => ada.getState().contacts.length === 1));
		const start = Date.now() + 3_600_000;
		const meeting = ada.createMeeting({
			title: 'Plan',
			start,
			end: start + 1_800_000,
			attendeeIds: ['bob'],
		});
		assert.match(meeting.chatId, /^g-/u);
		assert.ok(await until(() => bob.getState().chats.some((c) => c.id === meeting.chatId)));
		bob.select(meeting.chatId);
		assert.ok(await until(() => bob.getState().messages.some((m) => m.meetingId === meeting.id)));
	});

	it('refuses a case variant of a chat room', async () => {
		const workspace = 'chat-ws-6';
		const id = directChatId('ada', 'bob');
		const first = await dial(workspace, id, { user: 'ada', members: 'ada,bob' });
		assert.equal(first.status, 'open');
		const alias = chatRoomId(workspace, id).toUpperCase();
		const status = await new Promise((resolve) => {
			const ws = new WebSocket(`ws://127.0.0.1:${port}/sync/${alias}?user=eve`);
			sockets.push(ws);
			ws.on('open', () => resolve('open'));
			ws.on('unexpected-response', (_req, res) => resolve(res.statusCode));
			ws.on('error', () => resolve(0));
		});
		assert.equal(status, 403);
	});
});

describe('private chats with accounts', () => {
	let dir;
	let server;
	let port;
	const sockets = [];
	const sessions = {};
	const users = { ada: 'a-pw', bob: 'b-pw', eve: 'e-pw' };

	before(async () => {
		dir = mkdtempSync(join(tmpdir(), 'teams-chat-accounts-'));
		server = createTeamsServer({ dataDir: dir, users, log: () => {} });
		port = (await server.listen(0, '127.0.0.1')).port;
		for (const [user, credential] of Object.entries(users)) {
			const res = await fetch(`http://127.0.0.1:${port}/auth/session`, {
				method: 'POST',
				body: JSON.stringify({ user, credential }),
			});
			sessions[user] = (await res.json()).session;
		}
	});
	after(async () => {
		for (const ws of sockets) ws.terminate();
		await server.close();
		rmSync(dir, { recursive: true, force: true });
	});

	/** A socket that applies what the server sends to `doc` and sends what `doc` changes. */
	const connect = (room, user, query = '') =>
		new Promise((resolve) => {
			const doc = new Y.Doc();
			const ws = new WebSocket(`ws://127.0.0.1:${port}/sync/${room}${query}`, [
				'openteams',
				`openteams.session.${sessions[user]}`,
			]);
			sockets.push(ws);
			const peer = { ws, doc, status: 'connecting', closed: null };
			const send = (write) => {
				const encoder = encoding.createEncoder();
				encoding.writeVarUint(encoder, 0);
				write(encoder);
				if (ws.readyState === WebSocket.OPEN) ws.send(encoding.toUint8Array(encoder));
			};
			doc.on('update', (update, origin) => {
				if (origin !== peer) send((encoder) => syncProtocol.writeUpdate(encoder, update));
			});
			ws.on('message', (data) => {
				const decoder = decodingOf(data);
				if (decoder === null) return;
				const encoder = encoding.createEncoder();
				encoding.writeVarUint(encoder, 0);
				syncProtocol.readSyncMessage(decoder, encoder, doc, peer);
				if (encoding.length(encoder) > 1) ws.send(encoding.toUint8Array(encoder));
			});
			ws.on('open', () => {
				peer.status = 'open';
				send((encoder) => syncProtocol.writeSyncStep1(encoder, doc));
				resolve(peer);
			});
			ws.on('close', (code, reason) => (peer.closed = { code, reason: reason.toString() }));
			ws.on('unexpected-response', (_req, res) => {
				peer.status = res.statusCode;
				resolve(peer);
			});
			ws.on('error', () => {});
		});
	const members = (peer) => peer.doc.getMap('teams:chat-members');

	it('keeps a squatter out of a group room and lets the creator add people', async () => {
		const room = chatRoomId('acct-1', createGroupChatId('ada'));
		assert.equal((await connect(room, 'eve', '?members=eve,ada,bob')).status, 403);
		assert.equal((await connect(room, 'bob', '?members=ada,bob')).status, 403);
		const ada = await connect(room, 'ada');
		assert.equal(ada.status, 'open');
		assert.ok(await until(() => members(ada).get('ada') === true));
		members(ada).set('bob', true);
		assert.ok(await until(async () => (await connect(room, 'bob')).status === 'open'));
		assert.equal((await connect(room, 'eve')).status, 403);
	});

	it('refuses a false entry and keeps the member it was aimed at connected', async () => {
		const room = chatRoomId('acct-2', directChatId('ada', 'bob'));
		const ada = await connect(room, 'ada', '?members=ada,bob');
		const bob = await connect(room, 'bob');
		assert.ok(await until(() => members(ada).get('bob') === true));
		members(ada).set('bob', false);
		assert.ok(await until(() => ada.closed !== null));
		assert.match(ada.closed.reason, /update rejected: malformed-row/u);
		await sleep(50);
		assert.equal(bob.closed, null);
	});

	it('keeps a room closed once its last member has left', async () => {
		const room = chatRoomId('acct-3', createGroupChatId('ada'));
		const ada = await connect(room, 'ada');
		assert.ok(await until(() => members(ada).get('ada') === true));
		members(ada).delete('ada');
		assert.ok(await until(() => ada.closed !== null));
		assert.equal(ada.closed.reason, 'not a member');
		assert.equal((await connect(room, 'ada')).status, 403);
		assert.equal((await connect(room, 'eve')).status, 403);
	});

	it('merges concurrent group changes and undoes those of someone already removed', async () => {
		const chat = createGroupChatId('ada');
		const open = async (user) => {
			const peer = await connect('acct-5', user);
			assert.ok(await until(() => peer.doc.getMap('teams:chats').has(chat) || user === 'ada'));
			return peer;
		};
		const ada = await open('ada');
		const person = (id) => ({ id, name: id });
		ada.doc.transact(() => {
			const row = new Y.Map();
			row.set('kind', 'group');
			row.set('name', 'Crew');
			row.set('members', [person('ada'), person('bob')]);
			row.set('createdAt', 1);
			row.set('createdBy', 'ada');
			ada.doc.getMap('teams:chats').set(chat, row);
		});
		const bob = await open('bob');
		const rowOf = (peer) => peer.doc.getMap('teams:chats').get(chat);
		// Written at the same moment: neither has seen the other's change.
		rowOf(ada).set('member:cy', 'cy');
		rowOf(bob).set('member:bob', null);
		const settled = (peer) =>
			until(() => rowOf(peer).get('member:cy') === 'cy' && rowOf(peer).get('member:bob') === null);
		assert.ok(await settled(ada));
		assert.ok(await settled(bob));
		// bob is out now, and his client has not noticed.
		rowOf(bob).set('name', 'Mine');
		rowOf(bob).set('member:eve', 'eve');
		assert.ok(await until(() => rowOf(bob).get('name') === 'Crew'));
		assert.ok(await until(() => rowOf(ada).get('name') === 'Crew'));
		assert.equal(rowOf(ada).has('member:eve'), false);
		assert.equal(bob.closed, null, 'he stays connected');
		bob.doc.getMap('teams:roster').set('bob', new Y.Map());
		assert.ok(await until(() => ada.doc.getMap('teams:roster').has('bob')));
	});

	it('closes the socket of someone who rewrites a directory row they are not in', async () => {
		const ada = await connect('acct-4', 'ada');
		const row = new Y.Map();
		ada.doc.transact(() => {
			row.set('kind', 'direct');
			row.set('name', '');
			row.set('members', [
				{ id: 'ada', name: 'Ada' },
				{ id: 'bob', name: 'Bob' },
			]);
			row.set('createdAt', 1);
			row.set('createdBy', 'ada');
			ada.doc.getMap('teams:chats').set(directChatId('ada', 'bob'), row);
		});
		const eve = await connect('acct-4', 'eve');
		const seen = () => eve.doc.getMap('teams:chats').get(directChatId('ada', 'bob'));
		assert.ok(await until(() => seen() !== undefined));
		eve.doc.transact(() => {
			seen().set('kind', 'group');
			seen().set('members', [...seen().get('members'), { id: 'eve', name: 'Eve' }]);
		});
		assert.ok(await until(() => eve.closed !== null));
		assert.match(eve.closed.reason, /update rejected: not-member/u);
		assert.equal(ada.closed, null);
	});
});

function decodingOf(data) {
	const decoder = decoding.createDecoder(new Uint8Array(data));
	return decoding.readVarUint(decoder) === 0 ? decoder : null;
}
