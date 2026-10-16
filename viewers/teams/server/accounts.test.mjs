// node --test server/accounts.test.mjs
// Session authentication: a claimed id is replaced by the authenticated one.
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, before, describe, it } from 'node:test';
import { WebSocket } from 'ws';
import { directChatId, shortHash } from './chat-access.mjs';
import { createTeamsServer } from './index.mjs';

const open = (url) =>
	new Promise((resolve) => {
		const ws = new WebSocket(url);
		ws.on('open', () => resolve({ ok: true, ws }));
		ws.on('unexpected-response', (_req, res) => resolve({ ok: false, status: res.statusCode }));
		ws.on('error', () => {});
	});

describe('accounts', () => {
	const dirs = [];
	const servers = [];
	const start = async (options) => {
		const dir = mkdtempSync(join(tmpdir(), 'teams-accounts-'));
		dirs.push(dir);
		const server = createTeamsServer({ dataDir: dir, ...options });
		servers.push(server);
		return { server, port: (await server.listen(0, '127.0.0.1')).port };
	};
	after(async () => {
		for (const s of servers) await s.close();
		for (const d of dirs) rmSync(d, { recursive: true, force: true });
	});

	const login = async (port, user, credential) => {
		const res = await fetch(`http://127.0.0.1:${port}/auth/session`, {
			method: 'POST',
			body: JSON.stringify({ user, credential }),
		});
		return { status: res.status, body: res.status === 200 ? await res.json() : null };
	};
	const room = `chat-${shortHash('w')}-${directChatId('alice', 'bob')}`;
	const chatUrl = (port, query) => `ws://127.0.0.1:${port}/sync/${room}?${query}`;

	describe('with a users table', () => {
		let port;
		before(async () => {
			({ port } = await start({ users: { alice: 'a-pw', bob: 'b-pw', mallory: 'm-pw' } }));
		});

		it('rejects a wrong credential and an unknown user', async () => {
			assert.equal((await login(port, 'alice', 'nope')).status, 401);
			assert.equal((await login(port, 'nobody', 'a-pw')).status, 401);
			assert.equal((await login(port, 'alice')).status, 401);
		});

		it('refuses sockets without a valid session', async () => {
			assert.equal((await open(`ws://127.0.0.1:${port}/sync/general`)).status, 401);
			assert.equal((await open(`ws://127.0.0.1:${port}/sync/general?session=junk`)).status, 401);
			assert.equal((await open(`ws://127.0.0.1:${port}/signal/call`)).status, 401);
		});

		it('replaces a claimed id with the authenticated one', async () => {
			const alice = (await login(port, 'alice', 'a-pw')).body.session;
			const mallory = (await login(port, 'mallory', 'm-pw')).body.session;
			const members = 'alice,bob';
			// Mallory holds a valid session but claims to be alice: the session wins, so she is out.
			const forged = await open(chatUrl(port, `session=${mallory}&user=alice&members=${members}`));
			assert.equal(forged.status, 403);
			const real = await open(chatUrl(port, `session=${alice}&user=mallory&members=${members}`));
			assert.equal(real.ok, true);
			real.ws.close();
			const again = await open(chatUrl(port, `session=${mallory}&user=alice`));
			assert.equal(again.status, 403);
			const signal = await open(`ws://127.0.0.1:${port}/signal/call?session=${alice}`);
			assert.equal(signal.ok, true);
			signal.ws.close();
		});
	});

	describe('with an authenticate hook', () => {
		it('lets the provider decide the id and survives a throwing hook', async () => {
			const logs = [];
			const { port } = await start({
				log: (m) => logs.push(m),
				authenticate: ({ user, credential }) => {
					if (credential === 'boom') throw new Error('idp down');
					return credential === 'ok' ? { id: `${user}@corp` } : null;
				},
			});
			const ok = await login(port, 'carol', 'ok');
			assert.equal(ok.body.user, 'carol@corp');
			assert.equal((await login(port, 'carol', 'bad')).status, 401);
			assert.equal((await login(port, 'carol', 'boom')).status, 500);
			assert.ok(logs.every((m) => !m.includes('boom') && !m.includes(ok.body.session)));
			const sock = await open(`ws://127.0.0.1:${port}/sync/general?session=${ok.body.session}`);
			assert.equal(sock.ok, true);
			sock.ws.close();
		});
	});

	describe('dev default (no accounts)', () => {
		it('still works with zero config and issues sessions for claimed ids', async () => {
			const { port } = await start({});
			const general = await open(`ws://127.0.0.1:${port}/sync/general`);
			assert.equal(general.ok, true);
			general.ws.close();
			const claimed = await login(port, 'alice');
			assert.equal(claimed.body.user, 'alice');
			const chat = await open(chatUrl(port, 'user=alice&members=alice,bob'));
			assert.equal(chat.ok, true);
			chat.ws.close();
		});

		it('still checks the shared token on the session endpoint', async () => {
			const { port } = await start({ token: 'shared' });
			assert.equal((await login(port, 'alice')).status, 401);
			const res = await fetch(`http://127.0.0.1:${port}/auth/session`, {
				method: 'POST',
				headers: { Authorization: 'Bearer shared' },
				body: JSON.stringify({ user: 'alice' }),
			});
			assert.equal(res.status, 200);
		});

		it('rejects a forged or foreign session even in dev mode', async () => {
			const a = await start({ sessionSecret: 'one' });
			const b = await start({ sessionSecret: 'two' });
			const { body } = await login(a.port, 'alice');
			assert.equal(
				(await open(`ws://127.0.0.1:${b.port}/sync/general?session=${body.session}`)).status,
				401,
			);
		});
	});
});
