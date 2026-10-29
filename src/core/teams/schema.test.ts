import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { Transport } from '../collab/provider';
import {
	META_MAP,
	MIGRATIONS,
	MIGRATION_ORIGIN,
	SCHEMA_KEY,
	SCHEMA_VERSION,
	checkSchema,
	ensureSchema,
	readOnlyTransport,
	readSchemaVersion,
	watchSchema,
} from './schema';
import { createTeamsClient, type TeamsClient } from './store';
import { createTeamsWorkspace } from './workspace';

const tick = (ms = 0): Promise<void> => new Promise((r) => setTimeout(r, ms));

/** A workspace document as an older (unstamped) release wrote it: sparse rows. */
function legacyDoc(): Y.Doc {
	const doc = new Y.Doc();
	const channel = new Y.Map<unknown>();
	channel.set('name', 'General');
	channel.set('createdBy', 'ada');
	doc.getMap<Y.Map<unknown>>('teams:channels').set('general', channel);
	const row = new Y.Map<unknown>();
	row.set('authorId', 'ada');
	row.set('authorName', 'Ada');
	row.set('text', 'hello');
	doc.getMap<Y.Map<unknown>>('teams:messages:general').set('m1', row);
	return doc;
}

const stamp = (doc: Y.Doc, version: unknown): void => {
	doc.getMap<unknown>(META_MAP).set(SCHEMA_KEY, version);
};

describe('schema version', () => {
	it('pins the current version and keeps migrations contiguous', () => {
		expect(SCHEMA_VERSION).toBe(1);
		expect(MIGRATIONS.map((m) => m.version)).toEqual(
			Array.from({ length: SCHEMA_VERSION }, (_, i) => i + 1),
		);
	});

	it('stamps a fresh document once it may (after the first sync)', () => {
		const doc = new Y.Doc();
		expect(ensureSchema(doc, { stampEmpty: false })).toEqual({
			state: 'ok',
			version: null,
			supported: 1,
		});
		expect(readSchemaVersion(doc)).toBeNull();
		expect(ensureSchema(doc, { stampEmpty: true })).toMatchObject({ state: 'ok', version: 1 });
		expect(readSchemaVersion(doc)).toBe(1);
	});

	it('migrates a version 0 document, stamping it', () => {
		const doc = legacyDoc();
		expect(ensureSchema(doc, { stampEmpty: false })).toMatchObject({ state: 'ok', version: 1 });
		const channel = doc.getMap<Y.Map<unknown>>('teams:channels').get('general')!;
		expect(channel.toJSON()).toMatchObject({
			name: 'General',
			topic: '',
			kind: 'channel',
			archived: false,
			createdAt: 0,
		});
		const row = doc.getMap<Y.Map<unknown>>('teams:messages:general').get('m1')!;
		expect(row.toJSON()).toMatchObject({ text: 'hello', deleted: false, ts: 0 });
		// A message without mentions stays without mentions (valid in v1).
		expect(row.has('mentions')).toBe(false);
	});

	it('is idempotent and leaves a current document byte-identical', () => {
		const doc = legacyDoc();
		ensureSchema(doc, { stampEmpty: true });
		const before = Y.encodeStateAsUpdate(doc);
		ensureSchema(doc, { stampEmpty: true });
		ensureSchema(doc, { stampEmpty: true });
		expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
	});

	it('does not overwrite real values while migrating', () => {
		const doc = legacyDoc();
		doc.getMap<Y.Map<unknown>>('teams:channels').get('general')!.set('archived', true);
		doc.getMap<Y.Map<unknown>>('teams:messages:general').get('m1')!.set('deleted', true);
		ensureSchema(doc, { stampEmpty: true });
		expect(doc.getMap<Y.Map<unknown>>('teams:channels').get('general')!.get('archived')).toBe(true);
		expect(doc.getMap<Y.Map<unknown>>('teams:messages:general').get('m1')!.get('deleted')).toBe(
			true,
		);
	});

	it('tags the migration transaction with its origin', () => {
		const doc = legacyDoc();
		const origins: unknown[] = [];
		doc.on('update', (_u: Uint8Array, origin: unknown) => origins.push(origin));
		ensureSchema(doc, { stampEmpty: true });
		expect(origins).toEqual([MIGRATION_ORIGIN]);
	});

	it('refuses a newer document and leaves it untouched', () => {
		const doc = legacyDoc();
		stamp(doc, SCHEMA_VERSION + 1);
		const before = Y.encodeStateAsUpdate(doc);
		expect(ensureSchema(doc, { stampEmpty: true })).toEqual({
			state: 'unsupported-schema',
			found: SCHEMA_VERSION + 1,
			supported: SCHEMA_VERSION,
		});
		expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
		// Not even the v0 normalisation ran.
		expect(doc.getMap<Y.Map<unknown>>('teams:channels').get('general')!.has('topic')).toBe(false);
	});

	it('treats an unreadable stamp as unsupported', () => {
		for (const bad of ['1', 0, 1.5, -2, null]) {
			const doc = new Y.Doc();
			stamp(doc, bad);
			expect(checkSchema(doc)).toMatchObject({ state: 'unsupported-schema', found: null });
		}
	});

	it('reports a remote stamp through watchSchema', () => {
		const doc = new Y.Doc();
		const seen: string[] = [];
		const off = watchSchema(doc, (s) => seen.push(s.state));
		stamp(doc, 99);
		off();
		stamp(doc, 100);
		expect(seen).toEqual(['unsupported-schema']);
	});
});

describe('read-only transport', () => {
	const frame = (...parts: number[]): Uint8Array => {
		const e = encoding.createEncoder();
		for (const p of parts) encoding.writeVarUint(e, p);
		return encoding.toUint8Array(e);
	};
	const recorder = (): { transport: Transport; sent: Uint8Array[] } => {
		const sent: Uint8Array[] = [];
		return {
			sent,
			transport: { connect: vi.fn(), disconnect: vi.fn(), send: (d) => void sent.push(d) },
		};
	};

	it('passes everything while writable', () => {
		const { transport, sent } = recorder();
		const t = readOnlyTransport(transport, () => false);
		t.send(frame(0, 2));
		t.send(frame(0, 1));
		expect(sent).toHaveLength(2);
	});

	it('drops document changes but keeps reads and awareness while read-only', () => {
		const { transport, sent } = recorder();
		const t = readOnlyTransport(transport, () => true);
		t.send(frame(0, 0)); // sync step 1: asks for the state
		t.send(frame(1, 5)); // awareness
		t.send(frame(0, 1)); // sync step 2: carries our changes
		t.send(frame(0, 2)); // update
		expect(sent.map((d) => [d[0], d[1]])).toEqual([
			[0, 0],
			[1, 5],
		]);
	});
});

describe('workspace and client with a newer schema', () => {
	const clients: TeamsClient[] = [];
	afterEach(() => {
		for (const c of clients.splice(0)) c.destroy();
	});

	it('does not write to a newer document and says so', async () => {
		const doc = legacyDoc();
		stamp(doc, SCHEMA_VERSION + 1);
		const before = Y.encodeStateAsUpdate(doc);
		const ws = createTeamsWorkspace({
			workspaceId: 'schema-newer',
			user: { id: 'bob', name: 'Bob' },
			doc,
		});
		const seen: string[] = [];
		ws.schema.subscribe((s) => seen.push(s.state));
		await tick(3300); // past the sync grace, where a fresh workspace would seed and register
		expect(ws.schema.status().state).toBe('unsupported-schema');
		expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
		ws.destroy();
	}, 10_000);

	it('stamps a fresh workspace and registers after the first sync', async () => {
		const doc = new Y.Doc();
		const ws = createTeamsWorkspace({
			workspaceId: 'schema-fresh',
			user: { id: 'ann', name: 'Ann' },
			doc,
		});
		expect(readSchemaVersion(doc)).toBeNull();
		await tick(3300);
		expect(readSchemaVersion(doc)).toBe(1);
		expect(ws.chat.channels().map((c) => c.id)).toContain('general');
		expect(ws.roster.list().map((r) => r.id)).toContain('ann');
		ws.destroy();
	}, 10_000);

	it('migrates a stored v0 snapshot on open', () => {
		const stored = legacyDoc();
		const storage = new Map<string, string>();
		storage.set(
			'teams:doc:schema-v0',
			Buffer.from(Y.encodeStateAsUpdate(stored)).toString('base64'),
		);
		const client = createTeamsClient({
			workspaceId: 'schema-v0',
			user: { id: 'ada', name: 'Ada' },
			storage: {
				getItem: (k) => storage.get(k) ?? null,
				setItem: (k, v) => void storage.set(k, v),
			},
		});
		clients.push(client);
		expect(client.getState().schema).toMatchObject({ state: 'ok', version: 1 });
		expect(readSchemaVersion(client.workspace.session.doc)).toBe(1);
	});

	it('surfaces unsupported-schema in the client state and a notice', () => {
		const stored = legacyDoc();
		stamp(stored, 7);
		const storage = new Map<string, string>();
		storage.set(
			'teams:doc:schema-v7',
			Buffer.from(Y.encodeStateAsUpdate(stored)).toString('base64'),
		);
		const client = createTeamsClient({
			workspaceId: 'schema-v7',
			user: { id: 'ada', name: 'Ada' },
			storage: {
				getItem: (k) => storage.get(k) ?? null,
				setItem: (k, v) => void storage.set(k, v),
			},
		});
		clients.push(client);
		expect(client.getState().schema).toEqual({
			state: 'unsupported-schema',
			found: 7,
			supported: SCHEMA_VERSION,
		});
	});
});
