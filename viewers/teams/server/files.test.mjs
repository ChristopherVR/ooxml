// node --test server/files.test.mjs
// File permissions (private chat scopes), listing, deletion, partial-upload cleanup and quotas.
import assert from 'node:assert/strict';
import { existsSync, mkdirSync, mkdtempSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, describe, it } from 'node:test';
import { shortHash } from 'ooxml-core/teams';
import { WebSocket } from 'ws';
import { createFileStore, quotaKeyOf } from './files.mjs';
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

describe('file store bookkeeping', () => {
	it('keys quotas by workspace for both workspace and chat scopes', () => {
		assert.equal(quotaKeyOf('acme'), shortHash('acme'));
		assert.equal(quotaKeyOf(`chat-${shortHash('acme')}-g-1`), shortHash('acme'));
	});

	it('sweeps partial uploads and stale uploader records, keeping real files', () => {
		const dir = mkdtempSync(join(tmpdir(), 'teams-filestore-'));
		try {
			mkdirSync(join(dir, 'files', 'ws'), { recursive: true });
			writeFileSync(join(dir, 'files', 'ws', '.abc.part'), 'half');
			writeFileSync(join(dir, 'files', 'ws', 'keep.txt'), 'data');
			const store = createFileStore({ filesDir: join(dir, 'files'), metaDir: join(dir, 'meta') });
			assert.equal(store.bytesOf('ws'), 4, 'partial uploads do not count');
			store.added('ws', 'keep.txt', 0, 'ann');
			store.added('ws', 'gone.txt', 0, 'ann');
			assert.equal(store.sweep(60_000), 0, 'a fresh partial upload may still be in flight');
			assert.equal(store.sweep(0), 1);
			assert.deepEqual(readdirSync(join(dir, 'files', 'ws')), ['keep.txt']);
			assert.deepEqual(
				store.list('ws').map((f) => [f.name, f.uploader]),
				[['keep.txt', 'ann']],
			);
		} finally {
			rmSync(dir, { recursive: true, force: true });
		}
	});
});

describe('file permissions, lifecycle and quotas', () => {
	const dirs = [];
	const servers = [];
	const sockets = [];
	after(async () => {
		for (const s of sockets) s.terminate();
		for (const s of servers) await s.close();
		for (const d of dirs) rmSync(d, { recursive: true, force: true });
	});
	const start = async (options = {}) => {
		const dir = mkdtempSync(join(tmpdir(), 'teams-files-'));
		dirs.push(dir);
		const server = createTeamsServer({
			dataDir: dir,
			users: { ann: 'a', bob: 'b', eve: 'e', root: { credential: 'r', role: 'moderator' } },
			log: () => {},
			...options,
		});
		servers.push(server);
		const { port } = await server.listen(0, '127.0.0.1');
		const base = `http://127.0.0.1:${port}`;
		const sessions = {};
		const login = async (user, credential) => {
			const res = await fetch(`${base}/auth/session`, {
				method: 'POST',
				body: JSON.stringify({ user, credential }),
			});
			return (sessions[user] = (await res.json()).session);
		};
		const as = (user) => ({ 'X-Teams-Session': sessions[user] });
		await login('ann', 'a');
		await login('bob', 'b');
		await login('eve', 'e');
		await login('root', 'r');
		const req = (path, user, init = {}) =>
			fetch(`${base}${path}`, {
				...init,
				headers: { ...(user ? as(user) : {}), ...init.headers },
			});
		return { dir, port, base, sessions, as, req };
	};
	const chatRoom = `chat-${shortHash('acme')}-g-1`;
	const createChat = async (ctx) => {
		const ws = new WebSocket(`ws://127.0.0.1:${ctx.port}/sync/${chatRoom}?members=ann,bob`, [
			'openteams',
			`openteams.session.${ctx.sessions.ann}`,
		]);
		sockets.push(ws);
		await new Promise((ok, fail) => (ws.on('open', ok), ws.on('error', fail)));
		assert.ok(
			await until(async () => (await ctx.req(`/files/${chatRoom}`, 'ann')).status === 200),
			'the chat room exists',
		);
	};

	it('needs a signed session when identities are verified', async () => {
		const ctx = await start();
		assert.equal(
			(await fetch(`${ctx.base}/files/acme/a.txt`, { method: 'POST', body: 'x' })).status,
			401,
		);
		const bad = await fetch(`${ctx.base}/files/acme`, { headers: { 'X-Teams-Session': 'nope' } });
		assert.equal(bad.status, 401);
		assert.equal(
			(await ctx.req('/files/acme/a.txt', 'ann', { method: 'POST', body: 'x' })).status,
			201,
		);
	});

	it('limits private chat files to the members of that chat', async () => {
		const ctx = await start();
		await createChat(ctx);
		const up = await ctx.req(`/files/${chatRoom}/plan.txt`, 'ann', {
			method: 'POST',
			body: 'secret',
		});
		assert.equal(up.status, 201);
		const path = `/files/${chatRoom}/plan.txt`;
		assert.equal(await (await ctx.req(path, 'bob')).text(), 'secret');
		assert.equal((await ctx.req(path, 'eve')).status, 403, 'a workspace member outside the chat');
		assert.equal((await ctx.req(`/files/${chatRoom}`, 'eve')).status, 403);
		assert.equal(
			(await ctx.req(`/files/${chatRoom}/x.txt`, 'eve', { method: 'POST', body: 'x' })).status,
			403,
		);
		assert.equal(
			(await ctx.req(`/files/link/${chatRoom}/plan.txt`, 'eve', { method: 'POST' })).status,
			403,
		);
		assert.equal((await ctx.req(path, 'eve', { method: 'DELETE' })).status, 403);
		// A member gets a signed link that opens the file without a session.
		const link = await (
			await ctx.req(`/files/link/${chatRoom}/plan.txt`, 'bob', { method: 'POST' })
		).json();
		assert.equal(await (await fetch(`${ctx.base}${link.url}`)).text(), 'secret');
		// A room nobody created has no members, so nobody may use it.
		const unknown = `chat-${shortHash('acme')}-g-2`;
		assert.equal(
			(await ctx.req(`/files/${unknown}/a.txt`, 'ann', { method: 'POST', body: 'x' })).status,
			403,
		);
	});

	it('lets workspace members share channel files and lists them with their uploader', async () => {
		const ctx = await start();
		await ctx.req('/files/acme/a.txt', 'ann', { method: 'POST', body: 'one' });
		await ctx.req('/files/acme/b.txt', 'bob', { method: 'POST', body: 'two' });
		assert.equal(await (await ctx.req('/files/acme/a.txt', 'eve')).text(), 'one');
		const { files } = await (await ctx.req('/files/acme', 'eve')).json();
		assert.deepEqual(
			files.map((f) => [f.name, f.size, f.uploader]),
			[
				['a.txt', 3, 'ann'],
				['b.txt', 3, 'bob'],
			],
		);
	});

	it('deletes for the uploader or a moderator only', async () => {
		const ctx = await start();
		await ctx.req('/files/acme/a.txt', 'ann', { method: 'POST', body: 'one' });
		await ctx.req('/files/acme/b.txt', 'ann', { method: 'POST', body: 'two' });
		assert.equal((await ctx.req('/files/acme/a.txt', 'bob', { method: 'DELETE' })).status, 403);
		assert.equal((await ctx.req('/files/acme/a.txt', 'ann', { method: 'DELETE' })).status, 204);
		assert.equal((await ctx.req('/files/acme/a.txt', 'ann')).status, 404);
		assert.equal((await ctx.req('/files/acme/a.txt', 'ann', { method: 'DELETE' })).status, 404);
		assert.equal((await ctx.req('/files/acme/b.txt', 'root', { method: 'DELETE' })).status, 204);
		assert.deepEqual((await (await ctx.req('/files/acme', 'ann')).json()).files, []);
	});

	it('removes partial uploads left by a crash at start-up', async () => {
		const ctx = await start();
		await ctx.req('/files/acme/a.txt', 'ann', { method: 'POST', body: 'one' });
		const part = join(ctx.dir, 'files', 'acme', '.deadbeef.part');
		writeFileSync(part, 'torn');
		await servers.pop().close();
		const again = createTeamsServer({
			dataDir: ctx.dir,
			users: { ann: 'a' },
			log: () => {},
		});
		servers.push(again);
		assert.equal(existsSync(part), false);
		assert.equal(existsSync(join(ctx.dir, 'files', 'acme', 'a.txt')), true);
	});

	it('cleans up an upload that is cut short', async () => {
		const ctx = await start({ maxFileBytes: 10 });
		const res = await ctx.req('/files/acme/big.bin', 'ann', {
			method: 'POST',
			body: 'x'.repeat(50),
		});
		assert.equal(res.status, 413);
		await sleep(50);
		const stored = join(ctx.dir, 'files', 'acme');
		assert.deepEqual(existsSync(stored) ? readdirSync(stored) : [], []);
	});

	it('applies the quota per workspace, counting its chat rooms', async () => {
		const ctx = await start({ workspaceQuotaBytes: 10 });
		const ok = await ctx.req('/files/acme/a.txt', 'ann', { method: 'POST', body: '12345678' });
		assert.equal(ok.status, 201);
		assert.equal(
			(await ctx.req('/files/acme/b.txt', 'ann', { method: 'POST', body: '12345' })).status,
			507,
		);
		assert.equal(
			(await ctx.req('/files/other/b.txt', 'ann', { method: 'POST', body: '12345' })).status,
			201,
			'another workspace has its own allowance',
		);
		await ctx.req('/files/acme/a.txt', 'ann', { method: 'DELETE' });
		assert.equal(
			(await ctx.req('/files/acme/b.txt', 'ann', { method: 'POST', body: '12345' })).status,
			201,
			'deleting frees the quota',
		);
	});

	it('answers dev-mode chat files from the claimed user', async () => {
		const dir = mkdtempSync(join(tmpdir(), 'teams-files-dev-'));
		dirs.push(dir);
		const server = createTeamsServer({ dataDir: dir, insecureDev: true, log: () => {} });
		servers.push(server);
		const { port } = await server.listen(0, '127.0.0.1');
		const ws = new WebSocket(`ws://127.0.0.1:${port}/sync/${chatRoom}?user=ann&members=ann,bob`);
		sockets.push(ws);
		await new Promise((ok, fail) => (ws.on('open', ok), ws.on('error', fail)));
		const url = `http://127.0.0.1:${port}/files/${chatRoom}`;
		assert.ok(
			await until(
				async () => (await fetch(url, { headers: { 'X-Teams-User': 'ann' } })).status === 200,
			),
		);
		assert.equal((await fetch(url)).status, 403, 'no claimed user');
		assert.equal((await fetch(url, { headers: { 'X-Teams-User': 'eve' } })).status, 403);
	});
});
