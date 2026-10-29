// node --test server/hardening.test.mjs
// Abuse resistance of the reference server: bad input, default exposure, files and the limits.
import assert from 'node:assert/strict';
import { request as httpRequest } from 'node:http';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, describe, it } from 'node:test';
import { WebSocket } from 'ws';
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

describe('teams server hardening', () => {
	const dirs = [];
	const start = async (options = {}) => {
		const dir = mkdtempSync(join(tmpdir(), 'teams-server-hard-'));
		dirs.push(dir);
		const server = createTeamsServer({ dataDir: dir, insecureDev: true, ...options });
		const port = (await server.listen(0, '127.0.0.1')).port;
		return { server, port };
	};
	const raw = (port, path, method = 'GET') =>
		new Promise((resolve, reject) => {
			const req = httpRequest({ host: '127.0.0.1', port, path, method }, (res) => {
				res.resume();
				res.on('end', () => resolve(res.statusCode));
			});
			req.on('error', reject);
			req.end();
		});
	const rawUpgrade = (port, path) =>
		new Promise((resolve) => {
			const req = httpRequest({
				host: '127.0.0.1',
				port,
				path,
				headers: {
					Connection: 'Upgrade',
					Upgrade: 'websocket',
					'Sec-WebSocket-Version': '13',
					'Sec-WebSocket-Key': 'dGhlIHNhbXBsZSBub25jZQ==',
				},
			});
			req.on('response', (res) => resolve(res.statusCode));
			req.on('upgrade', () => resolve(101));
			req.on('error', () => resolve(0));
			req.end();
		});
	const open = (port, path, headers = {}) =>
		new Promise((resolve) => {
			const ws = new WebSocket(`ws://127.0.0.1:${port}${path}`, { headers });
			const got = [];
			const closed = new Promise((r) => ws.on('close', (code) => r(code)));
			ws.on('message', (d) => {
				try {
					got.push(JSON.parse(d.toString()));
				} catch {}
			});
			ws.on('open', () => resolve({ ok: true, ws, got, closed }));
			ws.on('error', () => resolve({ ok: false }));
			ws.on('unexpected-response', () => resolve({ ok: false }));
		});
	after(() => {
		for (const dir of dirs) rmSync(dir, { recursive: true, force: true });
	});

	it('answers bad percent-encoding with 400 instead of crashing', async () => {
		const { server, port } = await start();
		assert.equal(await rawUpgrade(port, '/sync/%E0%A4%A'), 400);
		assert.equal(await raw(port, '/files/ws1/%E0%A4%A', 'POST'), 400);
		assert.equal(await raw(port, '/files/ws1/%E0%A4%A'), 400);
		assert.equal(await raw(port, '/health'), 200, 'the process survived');
		await server.close();
	});

	it('refuses a non-loopback host without a token unless insecure', async () => {
		const dir = mkdtempSync(join(tmpdir(), 'teams-server-host-'));
		dirs.push(dir);
		const bare = createTeamsServer({ dataDir: dir, insecureDev: true });
		await assert.rejects(bare.listen(0, '0.0.0.0'), /refusing to listen/u);
		const withToken = createTeamsServer({ dataDir: dir, insecureDev: true, token: 't' });
		await withToken.listen(0, '0.0.0.0');
		await withToken.close();
		const insecure = createTeamsServer({ dataDir: dir, insecureDev: true, insecure: true });
		await insecure.listen(0, '0.0.0.0');
		await insecure.close();
	});

	it('only lets loopback origins and header-less clients into a token-less server', async () => {
		const { server, port } = await start();
		assert.equal((await open(port, '/sync/r1', { Origin: 'https://evil.example' })).ok, false);
		const local = await open(port, '/sync/r1', { Origin: 'http://localhost:5173' });
		assert.equal(local.ok, true);
		local.ws.close();
		const v6 = await open(port, '/signal/r1', { Origin: 'http://[::1]:8080' });
		assert.equal(v6.ok, true);
		v6.ws.close();
		const bare = await open(port, '/sync/r2');
		assert.equal(bare.ok, true);
		bare.ws.close();
		const evilPost = await fetch(`http://127.0.0.1:${port}/files/ws1/a.txt`, {
			method: 'POST',
			body: 'x',
			headers: { Origin: 'https://evil.example' },
		});
		assert.equal(evilPost.status, 403);
		await server.close();
	});

	it('rejects dot-only names, never overwrites, and answers 413 up front', async () => {
		const { server, port } = await start({ maxFileBytes: 100 });
		const base = `http://127.0.0.1:${port}/files/ws1`;
		assert.equal((await fetch(`${base}/...`, { method: 'POST', body: 'x' })).status, 400);
		assert.equal((await fetch(`${base}/....`)).status, 400);
		const first = await (await fetch(`${base}/a.txt`, { method: 'POST', body: 'one' })).json();
		const second = await (await fetch(`${base}/a.txt`, { method: 'POST', body: 'two' })).json();
		assert.equal(first.url, '/files/ws1/a.txt');
		assert.notEqual(second.url, first.url);
		assert.match(second.url, /^\/files\/ws1\/a-[0-9a-f]{6}\.txt$/u);
		assert.equal(await (await fetch(`http://127.0.0.1:${port}${first.url}`)).text(), 'one');
		assert.equal(await (await fetch(`http://127.0.0.1:${port}${second.url}`)).text(), 'two');
		const big = await fetch(`${base}/big.bin`, { method: 'POST', body: 'x'.repeat(5000) });
		assert.equal(big.status, 413);
		assert.equal((await fetch(`${base}/big.bin`)).status, 404, 'no partial file is left behind');
		await server.close();
	});

	it('enforces the files quota', async () => {
		const { server, port } = await start({ filesQuotaBytes: 20 });
		const base = `http://127.0.0.1:${port}/files/ws1`;
		const post = (name) => fetch(`${base}/${name}`, { method: 'POST', body: 'x'.repeat(15) });
		assert.equal((await post('a.bin')).status, 201);
		assert.equal((await post('b.bin')).status, 507);
		await server.close();
	});

	it('limits connections per address and rooms', async () => {
		const { server, port } = await start({ maxConnectionsPerIp: 2 });
		const a = await open(port, '/signal/r1');
		const b = await open(port, '/signal/r2');
		assert.equal(a.ok && b.ok, true);
		assert.equal((await open(port, '/signal/r1')).ok, false, 'third connection from this address');
		a.ws.close();
		b.ws.close();
		await sleep(150);
		const c = await open(port, '/signal/r1');
		assert.equal(c.ok, true, 'slots free up on close');
		c.ws.close();
		await server.close();
		const rooms = await start({ maxRooms: 1 });
		const r1 = await open(rooms.port, '/signal/one');
		assert.equal(r1.ok, true);
		assert.equal((await open(rooms.port, '/signal/two')).ok, false, 'second room refused');
		assert.equal((await open(rooms.port, '/signal/one')).ok, true, 'existing room still joinable');
		await rooms.server.close();
	});

	it('closes sockets that exceed the message rate', async () => {
		const { server, port } = await start({ messageRate: 1, messageBurst: 3 });
		const a = await open(port, '/signal/r1');
		for (let i = 0; i < 20; i++) a.ws.send(JSON.stringify({ type: 'hello', from: 'a' }));
		assert.equal(await a.closed, 1008);
		await server.close();
	});

	it('rejects sync updates past the document size limit', async () => {
		const { server, port } = await start({ maxDocBytes: 50 });
		const a = await open(port, '/sync/r1');
		const update = new Uint8Array(120);
		update[1] = 2; // MESSAGE_SYNC, then the update sub-message
		update[2] = 117;
		a.ws.send(update);
		assert.equal(await a.closed, 1009);
		await server.close();
	});

	it('routes addressed signals to one peer and refuses a claimed peer id', async () => {
		const { server, port } = await start();
		const a = await open(port, '/signal/r1');
		const b = await open(port, '/signal/r1');
		const c = await open(port, '/signal/r1');
		a.ws.send(JSON.stringify({ type: 'hello', from: 'a' }));
		b.ws.send(JSON.stringify({ type: 'hello', from: 'b' }));
		c.ws.send(JSON.stringify({ type: 'hello', from: 'c' }));
		assert.ok(await until(() => a.got.length >= 2));
		a.ws.send(JSON.stringify({ type: 'description', from: 'a', to: 'b' }));
		assert.ok(await until(() => b.got.some((m) => m.type === 'description')));
		await sleep(100);
		assert.equal(
			c.got.some((m) => m.type === 'description'),
			false,
			'an addressed message is not broadcast',
		);
		const dup = await open(port, '/signal/r1');
		dup.ws.send(JSON.stringify({ type: 'hello', from: 'a' }));
		assert.equal(await dup.closed, 1008);
		for (const x of [a, b, c]) x.ws.close();
		await server.close();
	});
});
