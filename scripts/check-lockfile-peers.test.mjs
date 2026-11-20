import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { test } from 'node:test';

import { parseLock } from './check-lockfile-drift.mjs';
import { emptyPeerSlots, keySegments } from './check-lockfile-peers.mjs';

const entry = (id, meta = {}) => [id, '', meta, 'sha512-x'];
const plugin = entry('@analogjs/vite-plugin-angular@2.9.7', {
	peerDependencies: { '@angular/build': '^22.0.0' },
	optionalPeers: ['@angular/build'],
});

test('splits lockfile keys into package names, scopes kept whole', () => {
	assert.deepEqual(keySegments('demo/@angular/build/chokidar'), [
		'demo',
		'@angular/build',
		'chokidar',
	]);
	assert.deepEqual(keySegments('typescript'), ['typescript']);
});

test('flags a peer the lockfile nests under a workspace but cannot look up from its dependent', () => {
	// The shape that aborted cold installs: the build package only under the workspace.
	const lock = {
		packages: {
			'@analogjs/vite-plugin-angular': plugin,
			'demo/@angular/build': entry('@angular/build@22.3.0'),
		},
	};
	assert.deepEqual(emptyPeerSlots(lock), [
		{
			key: '@analogjs/vite-plugin-angular',
			peer: '@angular/build',
			nested: ['demo/@angular/build'],
		},
	]);
});

test('accepts the peer once the lockfile can look it up, at the root or beside the dependent', () => {
	const hoisted = {
		packages: {
			'@analogjs/vite-plugin-angular': plugin,
			'@angular/build': entry('@angular/build@22.3.0'),
			'demo/@angular/build': entry('@angular/build@22.3.0'),
		},
	};
	const besides = {
		packages: {
			'demo/@analogjs/vite-plugin-angular': plugin,
			'demo/@angular/build': entry('@angular/build@22.3.0'),
		},
	};
	assert.deepEqual(emptyPeerSlots(hoisted), []);
	assert.deepEqual(emptyPeerSlots(besides), []);
});

test('ignores a peer that nothing in the lockfile provides', () => {
	assert.deepEqual(emptyPeerSlots({ packages: { '@analogjs/vite-plugin-angular': plugin } }), []);
});

test("the repository's bun.lock gives every nested peer a slot, so a cold install cannot abort", () => {
	const lock = parseLock(readFileSync(new URL('../bun.lock', import.meta.url), 'utf8'));
	assert.deepEqual(emptyPeerSlots(lock), []);
});
