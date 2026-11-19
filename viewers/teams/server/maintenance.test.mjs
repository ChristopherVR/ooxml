// node --test server/maintenance.test.mjs
// Room compaction (including a write cut short), the /metrics endpoint and the JSON logger.
import assert from 'node:assert/strict';
import {
	existsSync,
	mkdirSync,
	mkdtempSync,
	readdirSync,
	readFileSync,
	rmSync,
	writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, describe, it } from 'node:test';
import { WebSocket } from 'ws';
import * as Y from 'yjs';
import { directChatId } from './chat-access.mjs';
import { createTeamsServer } from './index.mjs';
import { createLogger, createMetrics } from './metrics.mjs';
import {
	compactSnapshotFile,
	compactUpdate,
	QUARANTINE_DIR,
	migrateStorageNames,
	nameFromStorage,
	storageName,
	sweepRoomTemps,
	writeFileAtomic,
} from './room-store.mjs';
import * as fsReal from 'node:fs';

const tmp = () => mkdtempSync(join(tmpdir(), 'teams-maint-'));

/** A snapshot built from many small updates, so that re-encoding it as one is smaller. */
function fragmentedSnapshot() {
	const doc = new Y.Doc();
	const text = doc.getText('t');
	const updates = [];
	doc.on('update', (u) => updates.push(u));
	for (let i = 0; i < 60; i++) text.insert(text.length, `line ${i}\n`);
	for (let i = 0; i < 30; i++) text.delete(0, 5);
	return { doc, text: text.toString(), concatenated: Y.mergeUpdates(updates) };
}

describe('room compaction', () => {
	it('re-encodes a document as one update with the same state', () => {
		const { doc, text } = fragmentedSnapshot();
		const compact = compactUpdate(doc);
		assert.ok(compact);
		const copy = new Y.Doc();
		Y.applyUpdate(copy, compact);
		assert.equal(copy.getText('t').toString(), text);
	});

	it('rewrites a bloated snapshot file smaller and leaves a compact one alone', () => {
		const dir = tmp();
		try {
			const { doc, text } = fragmentedSnapshot();
			const file = join(dir, 'r.bin');
			writeFileSync(file, Buffer.from(Y.encodeStateAsUpdate(doc)));
			const first = compactSnapshotFile(file);
			const second = compactSnapshotFile(file);
			if (first) assert.ok(first.after < first.before);
			assert.equal(second, null, 'already compact');
			const check = new Y.Doc();
			Y.applyUpdate(check, readFileSync(file));
			assert.equal(check.getText('t').toString(), text);
		} finally {
			rmSync(dir, { recursive: true, force: true });
		}
	});

	it('keeps the old snapshot and leaves no temp file when the write is cut short', () => {
		const dir = tmp();
		try {
			const file = join(dir, 'r.bin');
			writeFileSync(file, 'old');
			const cut = (name) => ({
				...fsReal,
				[name]: () => {
					throw new Error(`crash in ${name}`);
				},
			});
			for (const step of ['writeSync', 'fsyncSync', 'renameSync']) {
				assert.throws(() => writeFileAtomic(file, Buffer.from('new content'), cut(step)), /crash/u);
				assert.equal(readFileSync(file, 'utf8'), 'old', `after a crash in ${step}`);
				assert.deepEqual(readdirSync(dir), ['r.bin'], `no temp left after ${step}`);
			}
			writeFileAtomic(file, Buffer.from('new content'));
			assert.equal(readFileSync(file, 'utf8'), 'new content');
		} finally {
			rmSync(dir, { recursive: true, force: true });
		}
	});

	it('sweeps temp files from a process that died mid-write', () => {
		const dir = tmp();
		try {
			writeFileSync(join(dir, 'r.bin'), 'x');
			writeFileSync(join(dir, 'r.bin.0a1b2c3d.tmp'), 'torn');
			assert.equal(sweepRoomTemps(dir), 1);
			assert.deepEqual(readdirSync(dir), ['r.bin']);
		} finally {
			rmSync(dir, { recursive: true, force: true });
		}
	});

	it('compacts on load and on demand without changing what clients see', async () => {
		const dir = tmp();
		const servers = [];
		try {
			const { doc, text } = fragmentedSnapshot();
			mkdirSyncRooms(dir);
			const file = join(dir, 'rooms', `${storageName('team')}.bin`);
			const idle = join(dir, 'rooms', `${storageName('idle')}.bin`);
			// Pad the snapshot with a redundant copy so the compacted form is strictly smaller.
			const padded = Y.mergeUpdates([Y.encodeStateAsUpdate(doc), Y.encodeStateAsUpdate(doc)]);
			writeFileSync(idle, padded);
			writeFileSync(file, padded);
			const server = createTeamsServer({
				dataDir: dir,
				insecureDev: true,
				log: () => {},
				compactIntervalMs: 0,
			});
			servers.push(server);
			server.compactRooms();
			const check = new Y.Doc();
			Y.applyUpdate(check, readFileSync(idle));
			assert.equal(check.getText('t').toString(), text);
			assert.ok(readFileSync(idle).length <= padded.length);
			assert.equal(existsSync(file), true);
		} finally {
			for (const s of servers) await s.close();
			rmSync(dir, { recursive: true, force: true });
		}
	});
});

describe('storage names', () => {
	it('keeps every room name apart on any disk and can be read back', () => {
		for (const name of ['a', 'general', 'A-b_C', 'NUL', 'x'.repeat(128)]) {
			assert.equal(nameFromStorage(storageName(name)), name);
			assert.match(storageName(name), /^=[a-z2-7]+$/u);
		}
		assert.notEqual(storageName('Room'), storageName('room'));
		assert.ok(`${storageName('x'.repeat(128))}.bin.00000000.tmp`.length < 255);
		assert.equal(nameFromStorage('general'), null, 'a raw name from an older version');
		assert.equal(nameFromStorage('=A'), null);
		assert.equal(nameFromStorage('=b'), null, 'not what any name encodes to');
	});

	it('stores an upper-case spelling of a chat room as that room only on a case-insensitive disk', () => {
		const dir = tmp();
		try {
			const real = 'chat-0123456789abcdef-g1';
			const alias = real.toUpperCase().replace('0123456789ABCDEF', '0123456789ABCDEF');
			writeFileSync(join(dir, `${alias}.bin`), 'content');
			const caseInsensitive = existsSync(join(dir, `${real}.bin`));
			assert.equal(migrateStorageNames(dir, '.bin'), 1);
			assert.deepEqual(readdirSync(dir), [`${storageName(caseInsensitive ? real : alias)}.bin`]);
		} finally {
			rmSync(dir, { recursive: true, force: true });
		}
	});

	it('never lets an old entry become a chat id of the current format', () => {
		const hash = '0123456789abcdef';
		const direct = `dm-${'a'.repeat(32)}`;
		const group = `g-${'b'.repeat(32)}-${'c'.repeat(24)}`;
		for (const raw of [
			`chat-${hash}-${direct}`,
			`chat-${hash}-${group}`,
			`CHAT-${hash}-${direct}`.toUpperCase(),
			`chat-${hash.toUpperCase()}-${direct}`,
			`Chat-${hash}-${group}`,
		]) {
			const dir = tmp();
			try {
				writeFileSync(join(dir, `${raw}.bin`), 'planted');
				const logs = [];
				assert.equal(
					migrateStorageNames(dir, '.bin', (m) => logs.push(m)),
					0,
					raw,
				);
				assert.deepEqual(readdirSync(dir), [QUARANTINE_DIR], raw);
				assert.equal(readFileSync(join(dir, QUARANTINE_DIR, `${raw}.bin`), 'utf8'), 'planted');
				assert.match(logs.join(), /quarantined/u);
				assert.equal(migrateStorageNames(dir, '.bin'), 0, 'a second start changes nothing');
				assert.deepEqual(readdirSync(dir), [QUARANTINE_DIR], raw);
			} finally {
				rmSync(dir, { recursive: true, force: true });
			}
		}
	});

	it("keeps an old server's squatter out of the new chat it prepared, rooms and files alike", async () => {
		const dir = tmp();
		const hash = '0123456789abcdef';
		const dm = directChatId('ann', 'bob');
		const room = `chat-${hash}-${dm}`;
		const alias = room.toUpperCase();
		const planted = new Y.Doc();
		for (const user of ['eve', 'ann', 'bob']) planted.getMap('teams:chat-members').set(user, true);
		mkdirSync(join(dir, 'rooms'));
		mkdirSync(join(dir, 'files', alias), { recursive: true });
		mkdirSync(join(dir, 'file-meta'));
		writeFileSync(join(dir, 'rooms', `${alias}.bin`), Y.encodeStateAsUpdate(planted));
		writeFileSync(join(dir, 'files', alias, 'planted.txt'), 'planted');
		writeFileSync(
			join(dir, 'file-meta', `${alias}.json`),
			JSON.stringify({ 'planted.txt': 'eve' }),
		);
		const server = createTeamsServer({ dataDir: dir, insecureDev: true, log: () => {} });
		const { port } = await server.listen(0, '127.0.0.1');
		const dial = (user, query = '') =>
			new Promise((resolve) => {
				const ws = new WebSocket(`ws://127.0.0.1:${port}/sync/${room}?user=${user}${query}`);
				ws.on('open', () => (ws.close(), resolve('open')));
				ws.on('unexpected-response', (_q, res) => resolve(res.statusCode));
				ws.on('error', () => {});
			});
		try {
			assert.equal(await dial('eve'), 403);
			assert.equal(await dial('ann', '&members=ann,bob'), 'open');
			assert.equal(await dial('eve'), 403, 'the new room lists only its two people');
			const file = await fetch(`http://127.0.0.1:${port}/files/${room}/planted.txt`, {
				headers: { 'X-Teams-User': 'ann' },
			});
			assert.equal(file.status, 404);
			for (const area of ['rooms', 'files', 'file-meta'])
				assert.equal(readdirSync(join(dir, area, QUARANTINE_DIR)).length, 1, area);
		} finally {
			await server.close();
			rmSync(dir, { recursive: true, force: true });
		}
	});

	it('leaves an entry alone when its new name is taken, and says so', () => {
		const dir = tmp();
		try {
			writeFileSync(join(dir, 'team.bin'), 'old');
			writeFileSync(join(dir, `${storageName('team')}.bin`), 'new');
			const logs = [];
			assert.equal(
				migrateStorageNames(dir, '.bin', (m) => logs.push(m)),
				0,
			);
			assert.equal(readFileSync(join(dir, 'team.bin'), 'utf8'), 'old');
			assert.match(logs.join(), /could not migrate team\.bin/u);
		} finally {
			rmSync(dir, { recursive: true, force: true });
		}
	});

	it('migrates a data directory at start-up and keeps rooms, files and access', async () => {
		const dir = tmp();
		const chat = 'chat-0123456789abcdef-g-g-1-1';
		const snapshot = (fill) => {
			const doc = new Y.Doc();
			fill(doc);
			return Y.encodeStateAsUpdate(doc);
		};
		mkdirSync(join(dir, 'rooms'));
		mkdirSync(join(dir, 'files', 'acme'), { recursive: true });
		mkdirSync(join(dir, 'files', chat), { recursive: true });
		mkdirSync(join(dir, 'file-meta'));
		writeFileSync(
			join(dir, 'rooms', 'acme.bin'),
			snapshot((doc) => doc.getMap('m').set('k', 'kept')),
		);
		writeFileSync(
			join(dir, 'rooms', `${chat}.bin`),
			snapshot((doc) => doc.getMap('teams:chat-members').set('ann', true)),
		);
		writeFileSync(join(dir, 'files', 'acme', 'a.txt'), 'shared');
		writeFileSync(join(dir, 'files', chat, 'b.txt'), 'private');
		writeFileSync(join(dir, 'file-meta', 'acme.json'), JSON.stringify({ 'a.txt': 'ann' }));
		const boot = async () => {
			const server = createTeamsServer({ dataDir: dir, insecureDev: true, log: () => {} });
			const { port } = await server.listen(0, '127.0.0.1');
			return { server, base: `http://127.0.0.1:${port}`, port };
		};
		const get = (ctx, path, user) =>
			fetch(`${ctx.base}${path}`, { headers: { 'X-Teams-User': user } });
		const dial = (ctx, room, user) =>
			new Promise((resolve) => {
				const ws = new WebSocket(`ws://127.0.0.1:${ctx.port}/sync/${room}?user=${user}`);
				ws.on('open', () => (ws.close(), resolve('open')));
				ws.on('unexpected-response', (_q, res) => resolve(res.statusCode));
				ws.on('error', () => {});
			});
		try {
			// Twice: the second start finds nothing left to move and sees the same data.
			for (let run = 0; run < 2; run++) {
				const ctx = await boot();
				try {
					assert.deepEqual(
						readdirSync(join(dir, 'rooms')).sort(),
						[`${storageName('acme')}.bin`, `${storageName(chat)}.bin`].sort(),
					);
					assert.equal(await (await get(ctx, '/files/acme/a.txt', 'ann')).text(), 'shared');
					const listed = await (await get(ctx, '/files/acme', 'ann')).json();
					assert.equal(listed.files[0].uploader, 'ann');
					assert.equal(await (await get(ctx, `/files/${chat}/b.txt`, 'ann')).text(), 'private');
					assert.equal((await get(ctx, `/files/${chat}/b.txt`, 'eve')).status, 403);
					assert.equal(await dial(ctx, chat, 'ann'), 'open');
					assert.equal(await dial(ctx, chat, 'eve'), 403);
					assert.equal(ctx.server.isSyncRoomEmpty('acme'), false);
				} finally {
					await ctx.server.close();
				}
			}
		} finally {
			rmSync(dir, { recursive: true, force: true });
		}
	});
});

function mkdirSyncRooms(dir) {
	fsReal.mkdirSync(join(dir, 'rooms'), { recursive: true });
}

describe('metrics', () => {
	it('renders counters and gauges in the Prometheus text format', () => {
		const m = createMetrics();
		const hit = m.counter('x_total', 'Things.');
		hit({ code: 'a"b' });
		hit({ code: 'a"b' }, 2);
		m.counter('empty_total', 'None yet.');
		m.gauge('open', 'Open.', () => 3);
		m.gauge('by_kind', 'Kinds.', () => [{ labels: { kind: 'sync' }, value: 1 }]);
		const text = m.render();
		assert.match(text, /# TYPE x_total counter/u);
		assert.match(text, /x_total\{code="a\\"b"\} 3/u);
		assert.match(text, /^empty_total 0$/mu);
		assert.match(text, /^open 3$/mu);
		assert.match(text, /by_kind\{kind="sync"\} 1/u);
	});

	describe('endpoint', () => {
		const dirs = [];
		const servers = [];
		after(async () => {
			for (const s of servers) await s.close();
			for (const d of dirs) rmSync(d, { recursive: true, force: true });
		});
		const start = async (options) => {
			const dir = tmp();
			dirs.push(dir);
			const server = createTeamsServer({
				dataDir: dir,
				users: { ann: 'a' },
				log: () => {},
				...options,
			});
			servers.push(server);
			const { port } = await server.listen(0, '127.0.0.1');
			return `http://127.0.0.1:${port}`;
		};

		it('is off without a metrics token', async () => {
			const base = await start({});
			assert.equal((await fetch(`${base}/metrics`)).status, 404);
		});

		it('needs the token and reports rooms, connections, rejections and failures', async () => {
			const base = await start({ metricsToken: 'm3trics', sessionRate: { perUser: 1 } });
			assert.equal((await fetch(`${base}/metrics`)).status, 401);
			assert.equal(
				(await fetch(`${base}/metrics`, { headers: { Authorization: 'Bearer nope' } })).status,
				401,
			);
			// Wrong password (auth failure), then the second attempt for the same user is rate limited.
			const login = (credential) =>
				fetch(`${base}/auth/session`, {
					method: 'POST',
					body: JSON.stringify({ user: 'ann', credential }),
				});
			assert.equal((await login('wrong')).status, 401);
			assert.equal((await login('wrong')).status, 429);
			await fetch(`${base}/files/acme/a.txt`, { method: 'POST', body: 'x' });
			const ws = new WebSocket(`${base.replace('http', 'ws')}/sync/room-1`);
			ws.on('error', () => {});
			await sleep(100);
			const res = await fetch(`${base}/metrics`, { headers: { Authorization: 'Bearer m3trics' } });
			assert.equal(res.status, 200);
			assert.match(res.headers.get('content-type'), /^text\/plain/u);
			const text = await res.text();
			assert.match(text, /^teams_rooms_open 0$/mu);
			assert.match(text, /teams_connections\{kind="sync"\} 0/u);
			assert.match(text, /teams_stored_bytes\{kind="files"\} 0/u);
			assert.match(text, /teams_auth_failures_total\{kind="credentials"\} 1/u);
			assert.match(text, /teams_auth_failures_total\{kind="session"\} 2/u);
			assert.match(text, /teams_rate_limited_total\{scope="auth"\} 1/u);
			assert.match(text, /teams_auth_failures_total\{kind="metrics"\} 2/u);
		});
	});
});

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
describe('structured logs', () => {
	it('writes plain lines by default and one JSON object per line on request', () => {
		const out = [];
		createLogger('text', (l) => out.push(l))('hello');
		assert.deepEqual(out, ['hello']);
		const json = [];
		const log = createLogger(
			'json',
			(l) => json.push(l),
			() => Date.UTC(2026, 0, 2),
		);
		log('room x: ok');
		log('WARNING: careful');
		assert.deepEqual(
			json.map((l) => JSON.parse(l)),
			[
				{ time: '2026-01-02T00:00:00.000Z', level: 'info', msg: 'room x: ok' },
				{ time: '2026-01-02T00:00:00.000Z', level: 'warn', msg: 'WARNING: careful' },
			],
		);
	});
});
