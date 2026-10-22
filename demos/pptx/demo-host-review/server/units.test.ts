import { describe, expect, it } from 'vitest';

import { createDirectory, viewerRole } from './directory';
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
