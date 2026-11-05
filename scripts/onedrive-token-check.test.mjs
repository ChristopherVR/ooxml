import assert from 'node:assert/strict';
import { test } from 'node:test';

import { describeToken, normalizeToken } from './onedrive-token-check.mjs';

const jwt = (claims) =>
	['e30', Buffer.from(JSON.stringify(claims)).toString('base64url'), 'sig'].join('.');

test('normalizeToken drops whitespace, quotes and a Bearer prefix', () => {
	assert.equal(normalizeToken('  Bearer abc.def  \n'), 'abc.def');
	assert.equal(normalizeToken('"bearer xyz"'), 'xyz');
	assert.equal(normalizeToken(undefined), '');
});

test('describeToken recognises an authorization code', () => {
	assert.match(describeToken('M.C512_BAY.2.U.0b1c2d3e-aaaa').notes.join(), /authorization code/u);
});

test('describeToken reads audience, expiry and scopes without the token', () => {
	const now = 2_000_000_000;
	const notes = describeToken(
		jwt({ aud: 'https://management.azure.com', exp: now - 600, scp: 'User.Read' }),
		now,
	).notes.join('\n');
	assert.match(notes, /another resource/u);
	assert.match(notes, /expired 10 minutes ago/u);
	assert.match(notes, /no Files\.\* scope/u);
	assert.doesNotMatch(notes, /sig/u);
});

test('describeToken accepts a valid Graph token', () => {
	const token = jwt({
		aud: '00000003-0000-0000-c000-000000000000',
		exp: 4_000_000_000,
		scp: 'Files.ReadWrite User.Read',
	});
	assert.deepEqual(describeToken(token, 2_000_000_000).notes, []);
});
