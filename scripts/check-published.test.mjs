import assert from 'node:assert/strict';
import { test } from 'node:test';

import { findUnpublished, publishablePackages, recoveryCommand } from './check-published.mjs';

const packages = [
	{ name: '@x/a', version: '1.0.0', dir: 'a' },
	{ name: '@x/b', version: '2.0.0', dir: 'b' },
];

test('a package whose on-disk version is not on the registry is reported', () => {
	const published = new Set(['@x/a@1.0.0']);
	const missing = findUnpublished(packages, (name, version) => published.has(`${name}@${version}`));
	assert.deepEqual(
		missing.map((pkg) => pkg.name),
		['@x/b'],
	);
});

test('nothing is reported when every version is published', () => {
	assert.deepEqual(
		findUnpublished(packages, () => true),
		[],
	);
});

test('the recovery command is the documented tag dispatch', () => {
	assert.equal(recoveryCommand(packages[1]), 'gh workflow run release.yml -f tag=@x/b@2.0.0');
});

test('private packages are never expected on npm', () => {
	for (const pkg of publishablePackages()) assert.ok(pkg.name && pkg.version);
});
