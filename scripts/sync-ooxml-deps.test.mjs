import assert from 'node:assert/strict';
import { test } from 'node:test';

import { bumpManifest, isNewer, summarize } from './sync-ooxml-deps.mjs';

const latest = { 'ooxml-core': '0.15.0', 'ooxml-ui': '0.21.0' };

test('compares versions numerically, not as strings', () => {
	assert.equal(isNewer('0.15.0', '0.9.0'), true);
	assert.equal(isNewer('0.9.0', '0.15.0'), false);
	assert.equal(isNewer('1.0.0', '1.0.0'), false);
});

test('moves caret and tilde ranges in every dependency section', () => {
	const { manifest, changes } = bumpManifest(
		{
			dependencies: { 'ooxml-core': '^0.14.1', lit: '^3.0.0' },
			devDependencies: { 'ooxml-ui': '~0.20.0' },
			peerDependencies: { 'ooxml-core': '^0.8.0' },
		},
		latest,
	);
	assert.equal(manifest.dependencies['ooxml-core'], '^0.15.0');
	assert.equal(manifest.devDependencies['ooxml-ui'], '~0.21.0');
	assert.equal(manifest.peerDependencies['ooxml-core'], '^0.15.0');
	assert.equal(manifest.dependencies.lit, '^3.0.0');
	assert.equal(changes.length, 3);
});

test('leaves wildcard, file, workspace and newer ranges alone', () => {
	const input = {
		dependencies: {
			'ooxml-core': '*',
			'ooxml-ui': 'file:../ui',
		},
		devDependencies: { 'ooxml-core': 'workspace:*' },
		peerDependencies: { 'ooxml-core': '^0.16.0' },
	};
	const { manifest, changes } = bumpManifest(input, latest);
	assert.deepEqual(manifest, input);
	assert.deepEqual(changes, []);
});

test('does not mutate its input', () => {
	const input = { dependencies: { 'ooxml-core': '^0.1.0' } };
	bumpManifest(input, latest);
	assert.equal(input.dependencies['ooxml-core'], '^0.1.0');
});

test('summarizes each package once', () => {
	const { changes } = bumpManifest(
		{
			dependencies: { 'ooxml-core': '^0.1.0', 'ooxml-ui': '^0.1.0' },
			devDependencies: { 'ooxml-core': '^0.1.0' },
		},
		latest,
	);
	assert.equal(summarize(changes), 'ooxml-core 0.15.0 and ooxml-ui 0.21.0');
});
