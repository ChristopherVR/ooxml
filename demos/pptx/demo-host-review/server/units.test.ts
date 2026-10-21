import * as encoding from 'lib0/encoding';
import { describe, expect, it } from 'vitest';
import * as awarenessProtocol from 'y-protocols/awareness';
import * as syncProtocol from 'y-protocols/sync';
import * as Y from 'yjs';

import { createDirectory, viewerRole } from './directory';
import {
	MESSAGE_AWARENESS,
	MESSAGE_SYNC,
	mayProcess,
	messageKind,
	ownAwareness,
} from './relay-policy';
import { signToken, verifySignIn, verifyTicket } from './tickets';

const SECRET = 'test-secret';

describe('tickets', () => {
	it('round-trips a ticket and rejects tampering, expiry and the wrong kind', () => {
		const exp = Date.now() + 60_000;
		const ticket = signToken(SECRET, {
			account: 'ben',
			review: 'q3-launch',
			permission: 'review',
			exp,
		});
		expect(verifyTicket(SECRET, ticket)).toEqual({
			account: 'ben',
			review: 'q3-launch',
			permission: 'review',
			exp,
		});
		expect(verifyTicket('other', ticket)).toBeNull();
		expect(verifyTicket(SECRET, ticket, exp + 1)).toBeNull();
		const [body = '', signature = ''] = ticket.split('.');
		const forged = Buffer.from(
			JSON.stringify({ account: 'ben', review: 'q3-launch', permission: 'owner', exp }),
		).toString('base64url');
		expect(verifyTicket(SECRET, `${forged}.${signature}`)).toBeNull();
		expect(verifyTicket(SECRET, `${body}.${signature}.x`)).toBeNull();
		// A room ticket is not a sign-in, and a sign-in is not a room ticket.
		expect(verifySignIn(SECRET, ticket)).toBeNull();
		expect(verifyTicket(SECRET, signToken(SECRET, { account: 'ben', exp }))).toBeNull();
	});
});

describe('directory', () => {
	it('maps permissions to viewer roles and lets only the owner change access', () => {
		const directory = createDirectory(new Uint8Array());
		expect(viewerRole('owner')).toBe('owner');
		expect(viewerRole('review')).toBe('collaborator');
		expect(viewerRole('view')).toBe('viewer');
		expect(viewerRole('none')).toBeNull();
		expect(directory.setPermission('q3-launch', 'ben', 'dee', 'review')).toBe(false);
		expect(directory.setPermission('q3-launch', 'ana', 'ana', 'none')).toBe(false);
		expect(directory.setPermission('q3-launch', 'ana', 'ben', 'owner')).toBe(false);
		expect(directory.setPermission('q3-launch', 'ana', 'dee', 'review')).toBe(true);
		expect(directory.permission('q3-launch', 'dee')).toBe('review');
	});

	it('gives AI only to licensed accounts that may edit', () => {
		const directory = createDirectory(new Uint8Array());
		expect(directory.canUseAi('q3-launch', 'ben')).toBe(true);
		expect(directory.canUseAi('q3-launch', 'chen')).toBe(false);
		directory.setPermission('q3-launch', 'ana', 'ben', 'view');
		expect(directory.canUseAi('q3-launch', 'ben')).toBe(false);
	});

	it('keeps versions from the owner only', () => {
		const directory = createDirectory(new Uint8Array());
		expect(directory.addVersion('q3-launch', 'ben', new Uint8Array([1]))).toBeNull();
		expect(directory.addVersion('q3-launch', 'ana', new Uint8Array([1]))?.number).toBe(1);
		expect(directory.addVersion('q3-launch', 'ana', new Uint8Array([2]))?.number).toBe(2);
		expect(directory.review('q3-launch')?.versions.map((v) => v.number)).toEqual([2, 1]);
	});
});

describe('relay policy', () => {
	const doc = new Y.Doc();
	doc.getMap('m').set('k', 1);
	const sync = (write: (encoder: encoding.Encoder) => void): Uint8Array => {
		const encoder = encoding.createEncoder();
		encoding.writeVarUint(encoder, MESSAGE_SYNC);
		write(encoder);
		return encoding.toUint8Array(encoder);
	};

	it('lets read-only connections ask for the document but not change it', () => {
		const step1 = sync((e) => syncProtocol.writeSyncStep1(e, doc));
		const step2 = sync((e) => syncProtocol.writeSyncStep2(e, doc));
		const update = sync((e) => syncProtocol.writeUpdate(e, Y.encodeStateAsUpdate(doc)));
		expect([step1, step2, update].map(messageKind)).toEqual([
			'sync-step1',
			'sync-step2',
			'sync-update',
		]);
		expect(mayProcess('sync-step1', false)).toBe(true);
		expect(mayProcess('sync-step2', false)).toBe(false);
		expect(mayProcess('sync-update', false)).toBe(false);
		expect(mayProcess('sync-update', true)).toBe(true);
		expect(mayProcess('awareness', false)).toBe(true);
		expect(messageKind(new Uint8Array([9]))).toBe('invalid');
		expect(messageKind(new Uint8Array())).toBe('invalid');
		expect(mayProcess('invalid', true)).toBe(false);
	});

	it('keeps each awareness client id with the connection that first spoke for it', () => {
		const a = new awarenessProtocol.Awareness(new Y.Doc());
		const b = new awarenessProtocol.Awareness(new Y.Doc());
		a.setLocalState({ name: 'A' });
		b.setLocalState({ name: 'B' });
		const owners = new Map<number, string>();
		const fromA = awarenessProtocol.encodeAwarenessUpdate(a, [a.clientID]);
		expect(ownAwareness(fromA, 'a', owners)).not.toBeNull();
		// Connection b replaying a's entry (to clear or impersonate it) is dropped.
		expect(ownAwareness(fromA, 'b', owners)).toBeNull();
		const both = awarenessProtocol.encodeAwarenessUpdate(b, [b.clientID]);
		expect(ownAwareness(both, 'b', owners)).not.toBeNull();
		expect(owners.get(b.clientID)).toBe('b');
		expect(MESSAGE_AWARENESS).toBe(1);
		a.destroy();
		b.destroy();
	});
});
