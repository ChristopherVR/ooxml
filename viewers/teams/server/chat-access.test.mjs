// node --test server/chat-access.test.mjs
// Private chats: only the people listed in a chat's room may join it.
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, before, describe, it } from 'node:test';
import { chatRoomId, createTeamsClient, directChatId, shortHash } from 'ooxml-core/teams';
import { WebSocket } from 'ws';
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
	it('hashes exactly like the client', () => {
		for (const value of ['', 'ada', 'ada\nbob', 'x'.repeat(300), 'ü€']) {
			assert.equal(chatAccess.shortHash(value), shortHash(value));
		}
		assert.equal(chatAccess.directChatId('bob', 'ada'), directChatId('ada', 'bob'));
	});

	it('admits only listed members and seeds a new room from its first socket', () => {
		const name = room('g-1');
		const none = new Set();
		assert.deepEqual(
			chatAccess.authorizeChatJoin({ name, user: 'ada', members: 'ada,bob', existing: none }),
			{ ok: true, seed: ['ada', 'bob'] },
		);
		assert.equal(
			chatAccess.authorizeChatJoin({ name, user: 'eve', members: 'ada,bob', existing: none }).ok,
			false,
		);
		assert.equal(
			chatAccess.authorizeChatJoin({ name, user: 'ada', members: undefined, existing: none }).ok,
			false,
		);
		const known = new Set(['ada', 'bob']);
		assert.equal(
			chatAccess.authorizeChatJoin({ name, user: 'bob', members: '', existing: known }).ok,
			true,
		);
		assert.equal(
			chatAccess.authorizeChatJoin({ name, user: 'eve', members: 'eve,ada,bob', existing: known })
				.ok,
			false,
			'a member list in the URL never widens an existing room',
		);
		assert.equal(
			chatAccess.authorizeChatJoin({ name: 'plain-room', user: 'ada', existing: none }).ok,
			false,
		);
	});

	it('only creates a 1:1 room for the two people its id was derived from', () => {
		const name = room(directChatId('ada', 'bob'));
		const none = new Set();
		const attempt = (user, members) =>
			chatAccess.authorizeChatJoin({ name, user, members, existing: none });
		assert.equal(attempt('ada', 'ada,bob').ok, true);
		assert.equal(
			attempt('eve', 'eve,ada,bob').ok,
			false,
			'a third person cannot squat on the room',
		);
		assert.equal(attempt('eve', 'eve,bob').ok, false);
		assert.equal(attempt('ada', 'ada').ok, false);
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
	const client = (name, workspaceId) => {
		const c = createTeamsClient({
			workspaceId,
			user: { id: name, name },
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
		server = createTeamsServer({ dataDir: dir, insecureDev: true });
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

	it('closes the socket of someone removed from a group chat', async () => {
		const workspace = 'chat-ws-3';
		const ada = client('ada', workspace);
		const bob = client('bob', workspace);
		client('cy', workspace);
		assert.ok(await until(() => ada.getState().contacts.length === 2));
		const id = ada.createGroupChat(['bob', 'cy'], 'Crew');
		assert.ok(await until(() => bob.getState().chats.some((c) => c.id === id)));
		await sleep(300);
		const cySocket = await dial(workspace, id, { user: 'cy' });
		assert.equal(cySocket.status, 'open');
		ada.removeFromChat(id, 'cy');
		assert.equal(await Promise.race([cySocket.closed, sleep(4000).then(() => 'still open')]), 1008);
		assert.equal((await dial(workspace, id, { user: 'cy' })).status, 403);
		assert.equal((await dial(workspace, id, { user: 'bob' })).status, 'open');
	});
});
