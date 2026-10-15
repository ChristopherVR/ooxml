import { test } from 'node:test';
import assert from 'node:assert/strict';
import { authenticateFromUsers, createSessionSigner } from './auth.mjs';

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

test('users table checks the credential', () => {
	const auth = authenticateFromUsers({ alice: 'pw' });
	assert.deepEqual(auth({ user: 'alice', credential: 'pw' }), { id: 'alice' });
	assert.equal(auth({ user: 'alice', credential: 'no' }), null);
	assert.equal(auth({ user: 'bob', credential: 'pw' }), null);
	assert.equal(auth({ user: 'alice' }), null);
});
