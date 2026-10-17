import { test } from 'node:test';
import assert from 'node:assert/strict';
import { authenticateFromUsers, createSessionSigner, plainTextUsers } from './auth.mjs';
import { hashPassword, verifyPassword } from './passwords.mjs';
import { createSlidingLimiter } from './rate-limit.mjs';

test('a token verifies to the user it was issued for', () => {
	const s = createSessionSigner({ secret: 'k' });
	assert.equal(s.verify(s.issue('alice').token), 'alice');
});

test('forged, tampered and foreign tokens are rejected', () => {
	const s = createSessionSigner({ secret: 'k' });
	const other = createSessionSigner({ secret: 'other' });
	const [v, payload, sig] = s.issue('alice').token.split('.');
	const evil = Buffer.from(JSON.stringify({ u: 'bob', exp: Date.now() + 1e6 })).toString(
		'base64url',
	);
	assert.equal(s.verify(`${v}.${evil}.${sig}`), null);
	assert.equal(s.verify(other.issue('alice').token), null);
	assert.equal(s.verify(`${v}.${payload}`), null);
	assert.equal(s.verify(''), null);
	assert.equal(s.verify(undefined), null);
});

test('a token expires', () => {
	let t = 1000;
	const s = createSessionSigner({ secret: 'k', ttlMs: 10, now: () => t });
	const { token } = s.issue('alice');
	assert.equal(s.verify(token), 'alice');
	t = 1011;
	assert.equal(s.verify(token), null);
});

test('an invalid user id is not issued', () => {
	assert.throws(() => createSessionSigner().issue('a b'));
});

test('users table checks the credential', async () => {
	const auth = authenticateFromUsers({ alice: 'pw' });
	assert.deepEqual(await auth({ user: 'alice', credential: 'pw' }), { id: 'alice' });
	assert.equal(await auth({ user: 'alice', credential: 'no' }), null);
	assert.equal(await auth({ user: 'bob', credential: 'pw' }), null);
	assert.equal(await auth({ user: 'alice' }), null);
});

test('users table accepts hashed and plain credentials side by side', async () => {
	const auth = authenticateFromUsers({
		alice: await hashPassword('correct horse'),
		bob: 'plain',
	});
	assert.deepEqual(await auth({ user: 'alice', credential: 'correct horse' }), { id: 'alice' });
	assert.equal(await auth({ user: 'alice', credential: 'wrong' }), null);
	assert.equal(await auth({ user: 'alice', credential: 'scrypt$x' }), null);
	assert.deepEqual(await auth({ user: 'bob', credential: 'plain' }), { id: 'bob' });
});

test('a hash is salted, versioned and never equals the password', async () => {
	const a = await hashPassword('pw');
	const b = await hashPassword('pw');
	assert.notEqual(a, b);
	assert.ok(a.startsWith('scrypt$16384$8$1$'));
	assert.ok(!a.split('$').includes('pw'));
	assert.equal(await verifyPassword('pw', a), true);
	assert.equal(await verifyPassword('pw', 'scrypt$3$8$1$AA$AA'), false);
	assert.equal(await verifyPassword('pw', 'scrypt$16384$8$1$AA'), false);
	assert.equal(await verifyPassword('pw', 'plain'), false);
	await assert.rejects(() => hashPassword(''));
});

test('plainTextUsers lists only unhashed entries', async () => {
	assert.deepEqual(plainTextUsers({ a: 'x', b: await hashPassword('y') }), ['a']);
});

test('revoking a user invalidates earlier sessions only', () => {
	const seen = [];
	const s = createSessionSigner({ secret: 'k', onRevoke: (u, g) => seen.push([u, g]) });
	const old = s.issue('alice').token;
	const other = s.issue('bob').token;
	assert.equal(s.revoke('alice'), 1);
	assert.equal(s.verify(old), null);
	assert.equal(s.verify(other), 'bob');
	assert.equal(s.verify(s.issue('alice').token), 'alice');
	assert.deepEqual(seen, [['alice', 1]]);
});

test('a revocation seeded from storage survives a new signer', () => {
	const first = createSessionSigner({ secret: 'k' });
	const stale = first.issue('alice').token;
	const second = createSessionSigner({ secret: 'k', generations: { alice: 1 } });
	assert.equal(second.verify(stale), null);
	assert.equal(second.verify(second.issue('alice').token), 'alice');
});

test('sliding limiter refuses over the limit and recovers', () => {
	let t = 0;
	const l = createSlidingLimiter({ limit: 2, windowMs: 10000, now: () => t });
	assert.equal(l.check('a').ok, true);
	t = 1000;
	assert.equal(l.check('a').ok, true);
	t = 2000;
	const refused = l.check('a');
	assert.equal(refused.ok, false);
	assert.equal(refused.retryAfterSec, 8);
	assert.equal(l.check('b').ok, true);
	t = 10001;
	assert.equal(l.check('a').ok, true);
});

test('sliding limiter bounds the number of keys', () => {
	const l = createSlidingLimiter({ limit: 1, windowMs: 1000, maxKeys: 2, now: () => 0 });
	assert.equal(l.check('a').ok, true);
	assert.equal(l.check('b').ok, true);
	assert.equal(l.check('c').ok, false);
});
