// node --test server/fail-closed.test.mjs
// An update that cannot be checked or applied leaves no trace: nothing is stored, nobody else hears
// of it and the sender's connection closes. No update makes the validator fail on purpose, so the
// server is given a copy of it that fails on request.
import assert from 'node:assert/strict';
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { registerHooks } from 'node:module';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, afterEach, describe, it } from 'node:test';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import { createGroupChatId } from 'ooxml-core/teams';
import { WebSocket } from 'ws';
import * as syncProtocol from 'y-protocols/sync';
import * as Y from 'yjs';
import { faults } from './fail-closed.validator.mjs';
import { storageName } from './room-store.mjs';

registerHooks({
	resolve(specifier, context, next) {
		if (specifier === './validator.generated.mjs' && context.parentURL?.endsWith('/index.mjs'))
			return {
				url: new URL('./fail-closed.validator.mjs', import.meta.url).href,
				shortCircuit: true,
			};
		return next(specifier, context);
	},
});
const { createTeamsServer } = await import('./index.mjs');

const ROOM = 'workspace';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const until = async (fn, ms = 4000) => {
	const end = Date.now() + ms;
	while (Date.now() < end) {
		if (fn()) return true;
		await sleep(20);
	}
	return false;
};
const frame = (write) => {
	const encoder = encoding.createEncoder();
	encoding.writeVarUint(encoder, 0);
	write(encoder);
	return encoding.toUint8Array(encoder);
};
const rowOf = (peer, id) => peer.doc.getMap('teams:chats').get(id);

describe('an update that fails', () => {
	const dirs = [];
	const running = [];
	const sockets = [];
	afterEach(() => Object.assign(faults, { revert: false, validate: false }));
	after(async () => {
		for (const ws of sockets) ws.terminate();
		for (const server of running) await server.close();
		for (const dir of dirs) rmSync(dir, { recursive: true, force: true });
	});

	async function start(options) {
		const dir = mkdtempSync(join(tmpdir(), 'teams-fail-closed-'));
		dirs.push(dir);
		const server = createTeamsServer({
			dataDir: dir,
			sessionRate: { perIp: 1e5, perUser: 1e5 },
			...options,
		});
		running.push(server);
		const { port } = await server.listen(0, '127.0.0.1');
		const connect = async (user, doc = new Y.Doc()) => {
			const protocols = ['openteams'];
			if (options.users) {
				const response = await fetch(`http://127.0.0.1:${port}/auth/session`, {
					method: 'POST',
					body: JSON.stringify({ user, credential: `pw-${user}` }),
				});
				protocols.push(`openteams.session.${(await response.json()).session}`);
			}
			const ws = new WebSocket(`ws://127.0.0.1:${port}/sync/${ROOM}?user=${user}`, protocols);
			sockets.push(ws);
			ws.binaryType = 'arraybuffer';
			const peer = { ws, doc, closed: null };
			const onUpdate = (update, origin) => {
				if (origin !== peer && ws.readyState === WebSocket.OPEN)
					ws.send(frame((e) => syncProtocol.writeUpdate(e, update)));
			};
			doc.on('update', onUpdate);
			ws.on('close', (code, reason) => {
				doc.off('update', onUpdate);
				peer.closed = { code, reason: reason.toString() };
			});
			ws.on('message', (data) => {
				const decoder = decoding.createDecoder(new Uint8Array(data));
				if (decoding.readVarUint(decoder) !== 0) return;
				const encoder = encoding.createEncoder();
				encoding.writeVarUint(encoder, 0);
				syncProtocol.readSyncMessage(decoder, encoder, doc, peer);
				if (encoding.length(encoder) > 1) ws.send(encoding.toUint8Array(encoder));
			});
			await new Promise((resolve, reject) => {
				ws.once('open', resolve);
				ws.once('error', reject);
			});
			ws.send(frame((e) => syncProtocol.writeSyncStep1(e, doc)));
			return peer;
		};
		/** The room as the next start of the server would load it. */
		const stored = async () => {
			await server.close();
			running.splice(running.indexOf(server), 1);
			const doc = new Y.Doc();
			const file = join(dir, 'rooms', `${storageName(ROOM)}.bin`);
			if (existsSync(file)) Y.applyUpdate(doc, readFileSync(file));
			return doc;
		};
		return { connect, stored };
	}

	const accounts = () => ({
		users: {
			ann: 'pw-ann',
			cy: 'pw-cy',
			root: { credential: 'pw-root', role: 'moderator' },
		},
	});

	/** A group of root, ann and cy. Cy goes offline, is removed, and adds someone while offline. */
	async function staleAdd() {
		const context = await start(accounts());
		const root = await context.connect('root');
		const ann = await context.connect('ann');
		const cy = await context.connect('cy');
		const id = createGroupChatId('root');
		root.doc.transact(() => {
			const row = new Y.Map();
			row.set('kind', 'group');
			row.set('name', 'Crew');
			row.set(
				'members',
				['root', 'ann', 'cy'].map((m) => ({ id: m, name: m })),
			);
			row.set('createdAt', 1);
			row.set('createdBy', 'root');
			root.doc.getMap('teams:chats').set(id, row);
		});
		assert.ok(await until(() => rowOf(ann, id) && rowOf(cy, id)));
		cy.ws.close();
		assert.ok(await until(() => cy.closed));
		rowOf(root, id).set('member:cy', null);
		assert.ok(await until(() => rowOf(ann, id).get('member:cy') === null));
		rowOf(cy, id).set('member:eve', 'eve');
		const heard = [];
		ann.doc.on('update', () => heard.push(rowOf(ann, id).get('member:eve')));
		return { ...context, id, ann, cyDoc: cy.doc, heard };
	}

	it('control: an out-of-date add is undone in one step and its sender stays connected', async () => {
		const { connect, id, ann, cyDoc, heard } = await staleAdd();
		const before = faults.revertCalls;
		const cy = await connect('cy', cyDoc);
		assert.ok(await until(() => faults.revertCalls > before), 'the revert ran');
		assert.ok(await until(() => !rowOf(cy, id).has('member:eve')), 'the sender is corrected');
		await sleep(200);
		assert.equal(rowOf(ann, id).has('member:eve'), false);
		assert.deepEqual(heard.filter(Boolean), [], 'nobody saw the add, even for a moment');
		assert.equal(cy.closed, null);
	});

	it('keeps nothing and closes the sender when a revert fails', async () => {
		const { connect, stored, id, ann, cyDoc, heard } = await staleAdd();
		faults.revert = true;
		const cy = await connect('cy', cyDoc);
		assert.ok(await until(() => cy.closed), 'the sender is closed');
		assert.equal(cy.closed.code, 1008);
		assert.match(cy.closed.reason, /^update rejected/u);
		await sleep(100);
		assert.deepEqual(heard.filter(Boolean), [], 'no other client heard of it');
		assert.equal(rowOf(ann, id).has('member:eve'), false);
		const row = (await stored()).getMap('teams:chats').get(id);
		assert.equal(row.has('member:eve'), false, 'not stored');
		assert.equal(row.get('member:cy'), null);
	});

	it('keeps nothing and closes the sender when validation fails', async () => {
		const { connect, stored } = await start(accounts());
		const ann = await connect('ann');
		const root = await connect('root');
		await sleep(50);
		faults.validate = true;
		ann.doc.getMap('teams:roster').set('ann', new Y.Map());
		assert.ok(await until(() => ann.closed), 'the sender is closed');
		assert.equal(ann.closed.code, 1008);
		await sleep(100);
		assert.equal(root.doc.getMap('teams:roster').has('ann'), false);
		assert.equal(root.closed, null, 'other people stay connected');
		assert.equal((await stored()).getMap('teams:roster').has('ann'), false);
	});

	it('keeps nothing of an update that breaks off, on a socket nothing validates', async () => {
		const { connect, stored } = await start({ insecureDev: true });
		const ann = await connect('ann');
		const bob = await connect('bob');
		const source = new Y.Doc();
		source.getMap('notes').set('half', 'written');
		const whole = Y.encodeStateAsUpdate(source);
		ann.ws.send(frame((e) => syncProtocol.writeUpdate(e, whole.subarray(0, whole.length - 1))));
		assert.ok(await until(() => ann.closed), 'the sender is closed');
		assert.equal(ann.closed.code, 1008);
		await sleep(100);
		assert.equal(bob.doc.getMap('notes').has('half'), false);
		assert.equal((await stored()).getMap('notes').has('half'), false);
	});

	it('ignores what a refused socket sent after the update that was refused', async () => {
		const { connect, stored } = await start({ insecureDev: true });
		const ann = await connect('ann');
		const source = new Y.Doc();
		source.getMap('notes').set('later', 'written');
		const whole = Y.encodeStateAsUpdate(source);
		ann.ws.send(frame((e) => syncProtocol.writeUpdate(e, whole.subarray(0, whole.length - 1))));
		ann.ws.send(frame((e) => syncProtocol.writeUpdate(e, whole)));
		assert.ok(await until(() => ann.closed));
		assert.equal((await stored()).getMap('notes').has('later'), false);
	});
});
