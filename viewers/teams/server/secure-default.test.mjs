// node --test server/secure-default.test.mjs
// Secure by default (refuse to start without authentication), graceful flush on close, the
// opt-in trusted proxy address for the auth rate limiter and the session in a subprotocol.
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, describe, it } from 'node:test';
import * as encoding from 'lib0/encoding';
import * as syncProtocol from 'y-protocols/sync';
import * as Y from 'yjs';
import { WebSocket } from 'ws';
import { createTeamsServer, forwardedClient } from './index.mjs';

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const binPath = new URL('./bin.mjs', import.meta.url).pathname.replace(/^\/(\w:)/u, '$1');

describe('secure by default', () => {
	const dirs = [];
	const servers = [];
	const dir = () => {
		const d = mkdtempSync(join(tmpdir(), 'teams-secure-'));
		dirs.push(d);
		return d;
	};
	const start = async (options) => {
		const server = createTeamsServer({ dataDir: dir(), ...options });
		servers.push(server);
		return { server, port: (await server.listen(0, '127.0.0.1')).port };
	};
	after(async () => {
		for (const s of servers) await s.close();
		for (const d of dirs) rmSync(d, { recursive: true, force: true });
	});

	it('refuses to start with no accounts and no authenticate hook', () => {
		assert.throws(
			() => createTeamsServer({ dataDir: dir() }),
			(error) => /insecureDev/u.test(error.message) && /--insecure-dev/u.test(error.message),
		);
		// A shared token alone is not authentication: ids are still claims.
		assert.throws(() => createTeamsServer({ dataDir: dir(), token: 't' }), /refusing to start/u);
	});

	it('starts with insecureDev and warns loudly', async () => {
		const logs = [];
		const { port } = await start({ insecureDev: true, log: (m) => logs.push(m) });
		assert.equal((await fetch(`http://127.0.0.1:${port}/health`)).status, 200);
		assert.ok(logs.some((m) => m.includes('WARNING') && m.includes('INSECURE DEV')));
	});

	it('starts with accounts, an authenticate hook or authorizeSync, without a warning', async () => {
		for (const options of [
			{ users: { alice: 'pw' } },
			{ authenticate: () => null },
			{ authorizeSync: () => 'write' },
		]) {
			const logs = [];
			const { port } = await start({ ...options, log: (m) => logs.push(m) });
			assert.equal((await fetch(`http://127.0.0.1:${port}/health`)).status, 200);
			assert.equal(
				logs.some((m) => m.includes('INSECURE DEV')),
				false,
			);
		}
	});

	const runBin = (args, env = {}) =>
		new Promise((resolve) => {
			const port = String(20000 + Math.floor(Math.random() * 20000));
			const child = spawn(process.execPath, [binPath, ...args], {
				env: { ...process.env, PORT: port, HOST: '127.0.0.1', TEAMS_DATA: dir(), ...env },
				stdio: ['ignore', 'pipe', 'pipe'],
			});
			let out = '';
			const done = (code) => {
				clearTimeout(timer);
				child.kill();
				resolve({ code, out });
			};
			const timer = setTimeout(() => done('timeout'), 8000);
			const read = (chunk) => {
				out += chunk;
				if (out.includes('teams server on')) done('started');
			};
			child.stdout.on('data', read);
			child.stderr.on('data', read);
			child.on('exit', (code) => done(code));
		});

	it('the command exits non-zero without the flag and starts with it', async () => {
		const refused = await runBin([]);
		assert.equal(refused.code, 1);
		assert.match(refused.out, /--insecure-dev/u);
		const started = await runBin(['--insecure-dev']);
		assert.equal(started.code, 'started');
		assert.match(started.out, /INSECURE DEV/u);
	});

	it('close() persists a write made just before it', async () => {
		const dataDir = dir();
		const server = createTeamsServer({ dataDir, insecureDev: true });
		const { port } = await server.listen(0, '127.0.0.1');
		const ws = new WebSocket(`ws://127.0.0.1:${port}/sync/flush-room`);
		await new Promise((ok, fail) => (ws.once('open', ok), ws.once('error', fail)));
		const doc = new Y.Doc();
		doc.getMap('m').set('k', 'written-just-now');
		const encoder = encoding.createEncoder();
		encoding.writeVarUint(encoder, 0);
		syncProtocol.writeUpdate(encoder, Y.encodeStateAsUpdate(doc));
		ws.send(encoding.toUint8Array(encoder));
		await sleep(150); // delivered, well inside the 500 ms debounce
		const file = join(dataDir, 'rooms', 'flush-room.bin');
		assert.equal(existsSync(file), false, 'not yet written by the debounce');
		await server.close();
		const restored = new Y.Doc();
		Y.applyUpdate(restored, readFileSync(file));
		assert.equal(restored.getMap('m').get('k'), 'written-just-now');
	});
});

describe('trusted proxy', () => {
	const servers = [];
	const dirs = [];
	after(async () => {
		for (const s of servers) await s.close();
		for (const d of dirs) rmSync(d, { recursive: true, force: true });
	});
	const start = async (options) => {
		const d = mkdtempSync(join(tmpdir(), 'teams-proxy-'));
		dirs.push(d);
		const server = createTeamsServer({
			dataDir: d,
			insecureDev: true,
			sessionRate: { perIp: 2, perUser: 1000, windowMs: 60_000 },
			...options,
		});
		servers.push(server);
		return (await server.listen(0, '127.0.0.1')).port;
	};
	const login = async (port, user, forwarded) => {
		const res = await fetch(`http://127.0.0.1:${port}/auth/session`, {
			method: 'POST',
			headers: forwarded ? { 'x-forwarded-for': forwarded } : {},
			body: JSON.stringify({ user }),
		});
		return res.status;
	};

	it('picks the entry the trusted proxy appended', () => {
		const req = (xff) => ({ headers: { 'x-forwarded-for': xff } });
		assert.equal(forwardedClient(req('1.1.1.1, 2.2.2.2'), false), null);
		assert.equal(forwardedClient(req('9.9.9.9, 2.2.2.2'), 1), '2.2.2.2');
		assert.equal(forwardedClient(req('9.9.9.9, 3.3.3.3, 2.2.2.2'), 2), '3.3.3.3');
		assert.equal(forwardedClient(req('2.2.2.2'), 2), null);
		assert.equal(forwardedClient({ headers: {} }, 1), null);
	});

	it('ignores X-Forwarded-For unless trusted', async () => {
		const port = await start({});
		const codes = [];
		for (let i = 0; i < 4; i++) codes.push(await login(port, 'a', `10.0.0.${i}`));
		assert.deepEqual(codes, [200, 200, 429, 429]);
	});

	it('keys the limiter on the forwarded address when trusted', async () => {
		const port = await start({ trustProxy: true });
		// Different clients behind the proxy do not share a budget.
		for (let i = 0; i < 4; i++) assert.equal(await login(port, 'a', `10.0.0.${i}`), 200);
		// One client is limited, and a spoofed left-hand entry does not reset it.
		const codes = [];
		for (let i = 0; i < 4; i++) codes.push(await login(port, 'a', `spoof${i}, 10.9.9.9`));
		assert.deepEqual(codes, [200, 200, 429, 429]);
	});
});

describe('session in the subprotocol header', () => {
	const dirs = [];
	let server;
	after(async () => {
		await server?.close();
		for (const d of dirs) rmSync(d, { recursive: true, force: true });
	});

	it('accepts the session and echoes only the bare marker', async () => {
		const d = mkdtempSync(join(tmpdir(), 'teams-proto-'));
		dirs.push(d);
		server = createTeamsServer({ dataDir: d, users: { alice: 'pw' } });
		const { port } = await server.listen(0, '127.0.0.1');
		const res = await fetch(`http://127.0.0.1:${port}/auth/session`, {
			method: 'POST',
			body: JSON.stringify({ user: 'alice', credential: 'pw' }),
		});
		const { session } = await res.json();
		const connect = (protocols) =>
			new Promise((resolve) => {
				const ws = new WebSocket(`ws://127.0.0.1:${port}/sync/general`, protocols);
				ws.once('open', () => resolve({ ok: true, protocol: ws.protocol, ws }));
				ws.once('error', () => resolve({ ok: false }));
				ws.once('unexpected-response', () => resolve({ ok: false }));
			});
		const good = await connect(['openteams', `openteams.session.${session}`]);
		assert.equal(good.ok, true);
		assert.equal(good.protocol, 'openteams');
		good.ws.close();
		assert.equal((await connect(['openteams', 'openteams.session.junk'])).ok, false);
		assert.equal((await connect(['openteams'])).ok, false);
	});
});
