// node --test server/session-hardening.test.mjs
// /auth/session rate limits, hashed credentials, sign-out and admin revocation.
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, describe, it } from 'node:test';
import { WebSocket } from 'ws';
import { createTeamsServer } from './index.mjs';
import { hashPassword } from './passwords.mjs';

describe('session hardening', () => {
	const dirs = [];
	const servers = [];
	const logs = [];
	const start = async (options) => {
		const dir = options.dataDir ?? mkdtempSync(join(tmpdir(), 'teams-hardening-'));
		dirs.push(dir);
		const server = createTeamsServer({ dataDir: dir, insecureDev: true, log: (m) => logs.push(m), ...options });
		servers.push(server);
		return { server, dir, port: (await server.listen(0, '127.0.0.1')).port };
	};
	after(async () => {
		for (const s of servers) await s.close();
		for (const d of dirs) rmSync(d, { recursive: true, force: true });
	});
	const post = (port, path, body, headers = {}) =>
		fetch(`http://127.0.0.1:${port}${path}`, {
			method: 'POST',
			headers,
			body: JSON.stringify(body),
		});
	const login = async (port, user, credential) => {
		const res = await post(port, '/auth/session', { user, credential });
		return { status: res.status, res, body: res.status === 200 ? await res.json() : null };
	};
	const status = (url) =>
		new Promise((resolve) => {
			const ws = new WebSocket(url);
			ws.on('open', () => resolve({ ok: true, ws }));
			ws.on('unexpected-response', (_req, res) => resolve({ ok: false, status: res.statusCode }));
			ws.on('error', () => {});
		});

	it('accepts a hashed credential and warns about plain text ones', async () => {
		logs.length = 0;
		const { port } = await start({
			users: { alice: await hashPassword('s3cret-pass'), bob: 'plain-pw' },
		});
		assert.equal((await login(port, 'alice', 's3cret-pass')).status, 200);
		assert.equal((await login(port, 'alice', 'plain-pw')).status, 401);
		assert.equal((await login(port, 'bob', 'plain-pw')).status, 200);
		const warning = logs.find((m) => m.includes('plain text'));
		assert.ok(warning?.includes('1 user credential'));
		assert.ok(!logs.join('\n').includes('s3cret-pass'));
		assert.ok(!logs.join('\n').includes('plain-pw'));
	});

	it('does not warn when every credential is hashed', async () => {
		logs.length = 0;
		await start({ users: { alice: await hashPassword('x') } });
		assert.equal(
			logs.some((m) => m.includes('plain text')),
			false,
		);
	});

	it('limits attempts per user with 429 and Retry-After', async () => {
		const { port } = await start({
			users: { carol: 'pw' },
			sessionRate: { perUser: 3, perIp: 100, windowMs: 60_000 },
		});
		for (let i = 0; i < 3; i++) assert.equal((await login(port, 'carol', 'bad')).status, 401);
		const refused = await login(port, 'carol', 'pw');
		assert.equal(refused.status, 429);
		const wait = Number(refused.res.headers.get('retry-after'));
		assert.ok(wait >= 1 && wait <= 60);
		// Another user is unaffected.
		assert.equal((await login(port, 'dave', 'bad')).status, 401);
	});

	it('limits attempts per IP across users', async () => {
		const { port } = await start({
			users: { erin: 'pw' },
			sessionRate: { perUser: 100, perIp: 4, windowMs: 60_000 },
		});
		for (let i = 0; i < 4; i++) assert.equal((await login(port, `u${i}`, 'bad')).status, 401);
		const refused = await login(port, 'erin', 'pw');
		assert.equal(refused.status, 429);
		assert.ok(refused.res.headers.get('retry-after'));
	});

	it('sign-out revokes the user sessions and drops open sockets', async () => {
		const { port } = await start({
			users: { frank: 'pw' },
			sessionSecret: 'fixed',
		});
		const first = (await login(port, 'frank', 'pw')).body;
		const second = (await login(port, 'frank', 'pw')).body;
		const socket = await status(`ws://127.0.0.1:${port}/sync/general?session=${first.session}`);
		assert.equal(socket.ok, true);
		const closed = new Promise((resolve) => socket.ws.on('close', (code) => resolve(code)));
		assert.equal((await post(port, '/auth/check', { session: second.session })).status, 204);
		const out = await post(port, '/auth/signout', { session: first.session });
		assert.equal(out.status, 204);
		assert.equal(await closed, 1008);
		for (const s of [first, second])
			assert.equal(
				(await status(`ws://127.0.0.1:${port}/sync/general?session=${s.session}`)).status,
				401,
			);
		assert.equal((await post(port, '/auth/signout', { session: first.session })).status, 401);
		assert.equal((await post(port, '/auth/check', { session: second.session })).status, 401);
		assert.equal((await post(port, '/auth/check', { session: 'junk' })).status, 401);
		// A fresh sign-in works again.
		const again = (await login(port, 'frank', 'pw')).body;
		const ok = await status(`ws://127.0.0.1:${port}/sync/general?session=${again.session}`);
		assert.equal(ok.ok, true);
		ok.ws.close();
	});

	it('keeps a revocation across a restart with the same secret and data dir', async () => {
		const { server, port, dir } = await start({ users: { gina: 'pw' }, sessionSecret: 'keep' });
		const old = (await login(port, 'gina', 'pw')).body;
		server.revokeSessions('gina');
		await server.close();
		const next = await start({ users: { gina: 'pw' }, sessionSecret: 'keep', dataDir: dir });
		assert.equal(
			(await status(`ws://127.0.0.1:${next.port}/sync/general?session=${old.session}`)).status,
			401,
		);
		const fresh = (await login(next.port, 'gina', 'pw')).body;
		const ok = await status(`ws://127.0.0.1:${next.port}/sync/general?session=${fresh.session}`);
		assert.equal(ok.ok, true);
		ok.ws.close();
	});

	it('admin revoke needs the admin token and is off without one', async () => {
		const off = await start({ users: { hal: 'pw' } });
		assert.equal((await post(off.port, '/auth/revoke', { user: 'hal' })).status, 404);
		const { port } = await start({ users: { hal: 'pw' }, adminToken: 'admin-secret' });
		const s = (await login(port, 'hal', 'pw')).body;
		assert.equal((await post(port, '/auth/revoke', { user: 'hal' })).status, 401);
		assert.equal(
			(await post(port, '/auth/revoke', { user: 'hal' }, { Authorization: 'Bearer nope' })).status,
			401,
		);
		const done = await post(
			port,
			'/auth/revoke',
			{ user: 'hal' },
			{ Authorization: 'Bearer admin-secret' },
		);
		assert.equal(done.status, 204);
		assert.equal(
			(await status(`ws://127.0.0.1:${port}/sync/general?session=${s.session}`)).status,
			401,
		);
	});
});
