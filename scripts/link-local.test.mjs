import assert from 'node:assert/strict';
import { test } from 'node:test';

import { linkManifest, restoreManifest } from './link-local.mjs';

const manifest = {
	dependencies: { 'ooxml-core': '^0.15.0', lit: '^3.0.0' },
	peerDependencies: { 'ooxml-ui': '^0.23.0' },
};

test('links the known packages to file paths and records the ranges', () => {
	const { manifest: linked, saved } = linkManifest(manifest, '/v/packages/react', '/o');
	assert.match(linked.dependencies['ooxml-core'], /^file:.*o\/src\/core$/);
	assert.match(linked.peerDependencies['ooxml-ui'], /^file:.*o\/src\/ui$/);
	assert.equal(linked.dependencies.lit, '^3.0.0');
	assert.equal(saved.length, 2);
});

test('restores exactly what was linked', () => {
	const { manifest: linked, saved } = linkManifest(manifest, '/v/packages/react', '/o');
	assert.deepEqual(restoreManifest(linked, saved), manifest);
});

test('does not relink an already linked range', () => {
	const once = linkManifest(manifest, '/v/p', '/o').manifest;
	assert.equal(linkManifest(once, '/v/p', '/o').saved.length, 0);
});
