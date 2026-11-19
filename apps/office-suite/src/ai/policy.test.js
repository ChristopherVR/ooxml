import assert from 'node:assert/strict';
import { test } from 'node:test';

import { normalizePolicy } from './policy.js';

test('a token in the public configuration is ignored', () => {
	const policy = normalizePolicy({
		collaboration: { syncUrl: 'wss://sync.example', token: 'shared-room-secret' },
	});
	assert.equal(policy.collaboration.syncUrl, 'wss://sync.example');
	assert.equal('token' in policy.collaboration, false);
	assert.doesNotMatch(JSON.stringify(policy), /shared-room-secret/u);
});

test('the credentials endpoint must be https, and needs a sync server', () => {
	const collaboration = (credentialsUrl, syncUrl = 'wss://sync.example') =>
		normalizePolicy({ collaboration: { syncUrl, credentialsUrl } }).collaboration;
	assert.equal(
		collaboration('https://host.example/sync-token').credentialsUrl,
		'https://host.example/sync-token',
	);
	for (const refused of [
		'http://host.example/sync-token',
		'javascript:alert(1)',
		'https://user:pass@host.example/t',
		'not a url',
	])
		assert.equal(collaboration(refused).credentialsUrl, '', refused);
	assert.equal(collaboration('https://host.example/t', '').credentialsUrl, '');
	assert.equal(normalizePolicy({}).collaboration.credentialsUrl, '');
});
