// node --test server/authorize-sync.test.mjs
// Per-room permissions: `authorizeSync` admits, refuses or admits read-only, `closeSync` drops a
// principal's sockets, and `handleUpgrade` serves the same sockets from a host's own HTTP server.
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, describe, it } from 'node:test';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import { WebSocket } from 'ws';
import * as syncProtocol from 'y-protocols/sync';
import * as Y from 'yjs';
import { createTeamsServer, syncGrant } from './index.mjs';

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const until = async (fn, ms = 4000) => {
	const end = Date.now() + ms;
	while (Date.now() < end) {
		if (await fn()) return true;
		await sleep(25);
	}
	return false;
};

/** A minimal y-websocket client: syncs `doc` with the room and sends its local updates. */
function connect(url) {
	const doc = new Y.Doc();
	const ws = new WebSocket(url);
	ws.binaryType = 'arraybuffer';
	const frame = (write) => {
		const encoder = encoding.createEncoder();
		encoding.writeVarUint(encoder, 0);
		write(encoder);
		return encoding.toUint8Array(encoder);
	};
	ws.on('message', (data) => {
		const decoder = decoding.createDecoder(new Uint8Array(data));
		if (decoding.readVarUint(decoder) !== 0) return;
		const encoder = encoding.createEncoder();
		encoding.writeVarUint(encoder, 0);
		syncProtocol.readSyncMessage(decoder, encoder, doc, 'remote');
		if (encoding.length(encoder) > 1) ws.send(encoding.toUint8Array(encoder));
	});
	doc.on('update', (update, origin) => {
		if (origin !== 'remote' && ws.readyState === WebSocket.OPEN)
			ws.send(frame((e) => syncProtocol.writeUpdate(e, update)));
	});
	const opened = new Promise((resolve, reject) => {
		ws.on('open', () => {
			ws.send(frame((e) => syncProtocol.writeSyncStep1(e, doc)));
			resolve(true);
		});
		ws.on('error', () => reject(new Error('refused')));
	});
	const closed = new Promise((resolve) => ws.on('close', (code) => resolve(code)));
	return { doc, ws, opened, closed };
}

describe('authorizeSync', () => {
	const dirs = [];
	const servers = [];
	const sockets = [];
	after(async () => {
		for (const s of sockets) s.terminate();
		for (const s of servers) await s.close();
		for (const d of dirs) rmSync(d, { recursive: true, force: true });
	});
	const start = async (authorizeSync) => {
		const dir = mkdtempSync(join(tmpdir(), 'teams-authz-'));
		dirs.push(dir);
		const server = createTeamsServer({ dataDir: dir, authorizeSync });
		servers.push(server);
		const { port } = await server.listen(0, '127.0.0.1');
		return { server, url: (room, query) => `ws://127.0.0.1:${port}/sync/${room}?${query}` };
	};
	const open = (url) => {
		const client = connect(url);
		sockets.push(client.ws);
		return client;
	};

	it('normalises verdicts', () => {
		assert.deepEqual(syncGrant('write', 'ada'), { access: 'write', principal: 'ada' });
		assert.deepEqual(syncGrant({ access: 'read', principal: 'p1' }, 'ada'), {
			access: 'read',
			principal: 'p1',
		});
		assert.equal(syncGrant(null, 'ada'), null);
		assert.equal(syncGrant('owner', 'ada'), null);
		assert.equal(syncGrant({ access: 'admin' }, 'ada'), null);
	});

	it('refuses, admits read-only and admits writers, and can decide asynchronously', async () => {
		const access = { writer: 'write', reader: 'read' };
		const { server, url } = await start(async ({ room, url: request }) => {
			await sleep(10);
			const who = request.searchParams.get('ticket');
			return room === 'deck-1' && access[who] ? { access: access[who], principal: who } : null;
		});
		await assert.rejects(open(url('deck-1', 'ticket=stranger')).opened, /refused/);
		await assert.rejects(open(url('deck-2', 'ticket=writer')).opened, /refused/);
		const writer = open(url('deck-1', 'ticket=writer'));
		const reader = open(url('deck-1', 'ticket=reader'));
		await Promise.all([writer.opened, reader.opened]);
		writer.doc.getMap('m').set('title', 'from writer');
		assert.ok(await until(() => reader.doc.getMap('m').get('title') === 'from writer'));
		reader.doc.getMap('m').set('title', 'from reader');
		await sleep(200);
		assert.equal(writer.doc.getMap('m').get('title'), 'from writer');
		const late = open(url('deck-1', 'ticket=writer'));
		await late.opened;
		assert.ok(await until(() => late.doc.getMap('m').get('title') === 'from writer'));
		assert.deepEqual(server.syncPrincipals('deck-1').sort(), ['reader', 'writer']);
		assert.equal(server.isSyncRoomEmpty('deck-1'), false);
		assert.equal(server.isSyncRoomEmpty('deck-9'), true);
	});

	it('closes only the named principal with closeSync', async () => {
		const { server, url } = await start(({ url: request }) => ({
			access: 'write',
			principal: request.searchParams.get('ticket'),
		}));
		const ada = open(url('deck-3', 'ticket=ada'));
		const bob = open(url('deck-3', 'ticket=bob'));
		await Promise.all([ada.opened, bob.opened]);
		assert.equal(server.closeSync('deck-3', 'bob'), 1);
		assert.equal(await bob.closed, 4403);
		await sleep(50);
		assert.equal(ada.ws.readyState, WebSocket.OPEN);
	});

	it('answers 500 when the hook throws', async () => {
		const { url } = await start(() => {
			throw new Error('directory down');
		});
		await assert.rejects(open(url('deck-4', 'ticket=x')).opened, /refused/);
	});

	it('serves sync sockets from a host server through handleUpgrade', async () => {
		const dir = mkdtempSync(join(tmpdir(), 'teams-authz-'));
		dirs.push(dir);
		const teams = createTeamsServer({ dataDir: dir, authorizeSync: () => 'write' });
		servers.push(teams);
		const host = createServer((_req, res) => res.end('host'));
		host.on('upgrade', (req, socket, head) => teams.handleUpgrade(req, socket, head));
		await new Promise((resolve) => host.listen(0, '127.0.0.1', resolve));
		servers.push({ close: () => new Promise((resolve) => host.close(resolve)) });
		const client = open(`ws://127.0.0.1:${host.address().port}/sync/hosted`);
		assert.equal(await client.opened, true);
	});
});
