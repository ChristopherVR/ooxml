// node --test server/authorship.test.mjs
// Server-side enforcement: authorship of messages (with a moderator role carried by the signed
// session) and the protocol version a client announces.
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, describe, it } from 'node:test';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import { WebSocket } from 'ws';
import * as syncProtocol from 'y-protocols/sync';
import * as Y from 'yjs';
import { authenticateFromUsers, createSessionSigner } from './auth.mjs';
import { createTeamsServer } from './index.mjs';

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const until = async (fn, ms = 3000) => {
	const end = Date.now() + ms;
	while (Date.now() < end) {
		if (await fn()) return true;
		await sleep(25);
	}
	return false;
};

const MESSAGES = 'teams:messages:general';

/** A minimal y-websocket client that sends its session in the WebSocket subprotocol. */
function connect(url, session) {
	const doc = new Y.Doc();
	const ws = new WebSocket(
		url,
		session ? ['openteams', `openteams.session.${session}`] : undefined,
	);
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
	const closed = new Promise((resolve) =>
		ws.on('close', (code, reason) => resolve({ code, reason: reason.toString() })),
	);
	return { doc, ws, opened, closed };
}

const row = (authorId, text = 'hi') => {
	const m = new Y.Map();
	m.set('authorId', authorId);
	m.set('text', text);
	m.set('deleted', false);
	return m;
};

describe('session roles', () => {
	it('carries the moderator role in the signed token', () => {
		const signer = createSessionSigner({ secret: 's' });
		const plain = signer.issue('ann');
		const mod = signer.issue('root', 'moderator');
		assert.deepEqual(signer.inspect(plain.token), { user: 'ann', role: undefined });
		assert.deepEqual(signer.inspect(mod.token), { user: 'root', role: 'moderator' });
		assert.equal(signer.verify(mod.token), 'root');
		assert.equal(signer.issue('x', 'admin').token.split('.').length, 3);
		assert.equal(signer.inspect(signer.issue('x', 'admin').token).role, undefined);
	});

	it('reads the role from a users entry', async () => {
		const authenticate = authenticateFromUsers({
			ann: 'a',
			root: { credential: 'r', role: 'moderator' },
		});
		assert.deepEqual(await authenticate({ user: 'ann', credential: 'a' }), { id: 'ann' });
		assert.deepEqual(await authenticate({ user: 'root', credential: 'r' }), {
			id: 'root',
			role: 'moderator',
		});
		assert.equal(await authenticate({ user: 'root', credential: 'x' }), null);
	});
});

describe('authorship and protocol checks', () => {
	const dirs = [];
	const servers = [];
	const sockets = [];
	const logs = [];
	after(async () => {
		for (const s of sockets) s.terminate();
		for (const s of servers) await s.close();
		for (const d of dirs) rmSync(d, { recursive: true, force: true });
	});
	const start = async (options = {}) => {
		const dir = mkdtempSync(join(tmpdir(), 'teams-authorship-'));
		dirs.push(dir);
		const server = createTeamsServer({
			dataDir: dir,
			users: { ann: 'a', bob: 'b', root: { credential: 'r', role: 'moderator' } },
			log: (m) => logs.push(m),
			...options,
		});
		servers.push(server);
		const { port } = await server.listen(0, '127.0.0.1');
		const login = async (user, credential) => {
			const res = await fetch(`http://127.0.0.1:${port}/auth/session`, {
				method: 'POST',
				body: JSON.stringify({ user, credential }),
			});
			return res.json();
		};
		const open = async (user, credential, query = 'proto=1', room = 'team') => {
			const { session } = await login(user, credential);
			const client = connect(`ws://127.0.0.1:${port}/sync/${room}?${query}`, session);
			sockets.push(client.ws);
			await client.opened;
			return client;
		};
		return { port, login, open };
	};

	it('reports the role in the session response', async () => {
		const { login } = await start();
		assert.equal((await login('root', 'r')).role, 'moderator');
		assert.equal((await login('ann', 'a')).role, undefined);
	});

	it('applies honest writes and closes a forged author with 1008', async () => {
		const { open } = await start();
		const ann = await open('ann', 'a');
		const bob = await open('bob', 'b');
		bob.doc.getMap(MESSAGES).set('m1', row('bob'));
		assert.ok(await until(() => ann.doc.getMap(MESSAGES).has('m1')));
		bob.doc.getMap(MESSAGES).set('m2', row('ann', 'forged'));
		const closed = await bob.closed;
		assert.equal(closed.code, 1008);
		assert.match(closed.reason, /forged-author/);
		await sleep(100);
		assert.equal(ann.doc.getMap(MESSAGES).has('m2'), false);
	});

	it('lets a moderator hide a message and refuses an ordinary member', async () => {
		const { open } = await start();
		const ann = await open('ann', 'a', 'proto=1', 'team2');
		ann.doc.getMap(MESSAGES).set('m1', row('ann'));
		const root = await open('root', 'r', 'proto=1', 'team2');
		assert.ok(await until(() => root.doc.getMap(MESSAGES).has('m1')));
		root.doc.getMap(MESSAGES).get('m1').set('deleted', true);
		assert.ok(await until(() => ann.doc.getMap(MESSAGES).get('m1').get('deleted') === true));
		ann.doc.getMap(MESSAGES).get('m1').set('deleted', false);
		const bob = await open('bob', 'b', 'proto=1', 'team2');
		assert.ok(await until(() => bob.doc.getMap(MESSAGES).has('m1')));
		bob.doc.getMap(MESSAGES).get('m1').set('text', 'defaced');
		assert.equal((await bob.closed).code, 1008);
	});

	it('does not enforce when enforceAuthorship is off', async () => {
		const { open } = await start({ enforceAuthorship: false });
		const ann = await open('ann', 'a', 'proto=1', 'team3');
		const bob = await open('bob', 'b', 'proto=1', 'team3');
		bob.doc.getMap(MESSAGES).set('m1', row('ann', 'forged'));
		assert.ok(await until(() => ann.doc.getMap(MESSAGES).has('m1')));
	});

	it('refuses an unsupported protocol version with a reason', async () => {
		const { open } = await start();
		const client = await open('ann', 'a', 'proto=2', 'team4');
		const closed = await client.closed;
		assert.equal(closed.code, 1008);
		assert.match(closed.reason, /unsupported protocol version 2.*speaks 1/);
		const junk = await open('ann', 'a', 'proto=abc', 'team4');
		assert.equal((await junk.closed).code, 1008);
	});

	it('accepts a client that announces no version, and logs it', async () => {
		const { open } = await start();
		logs.length = 0;
		const ann = await open('ann', 'a', '', 'team5');
		const bob = await open('bob', 'b', 'proto=1', 'team5');
		ann.doc.getMap(MESSAGES).set('m1', row('ann'));
		assert.ok(await until(() => bob.doc.getMap(MESSAGES).has('m1')));
		assert.ok(logs.some((l) => /announced no protocol version/.test(l)));
	});
});
