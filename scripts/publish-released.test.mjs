import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, test } from 'node:test';

import { PACKAGES } from './release-plan.mjs';
import { publishManifest, resolveTargets, verifyManifest } from './publish-released.mjs';

const [key, meta] = Object.entries(PACKAGES)[0];
const version = JSON.parse(
	await (await import('node:fs/promises')).readFile(`${meta.dir}/package.json`, 'utf8'),
).version;

test('a tag resolves to exactly one package, scoped names included', () => {
	const [target] = resolveTargets({ tag: `${meta.npm}@1.2.3` });
	assert.deepEqual(target, { key, npm: meta.npm, dir: meta.dir, version: '1.2.3' });
});

test('a malformed or unknown tag is rejected', () => {
	assert.throws(() => resolveTargets({ tag: 'v1.2.3' }), /Invalid tag|Unknown package/);
	assert.throws(() => resolveTargets({ tag: 'nope@1.0.0' }), /Unknown package/);
	assert.throws(() => resolveTargets({ tag: `${meta.npm}@1.0.0; rm -rf /` }), /Invalid tag/);
});

test('a plan resolves to its released packages in plan order', () => {
	const plan = {
		order: ['b', 'a', 'c'],
		packages: {
			a: { release: true, npm: '@x/a', dir: 'packages/a', version: '1.0.0' },
			b: { release: false, npm: '@x/b', dir: 'packages/b', version: '1.0.0' },
			c: { release: true, npm: '@x/c', dir: 'packages/c', version: '2.0.0' },
		},
	};
	assert.deepEqual(
		resolveTargets({ plan }).map((t) => t.npm),
		['@x/a', '@x/c'],
	);
});

test('the manifest on disk must be the version being published', () => {
	assert.throws(
		() => verifyManifest({ npm: meta.npm, dir: meta.dir, version: '9.9.9' }),
		/on disk/,
	);
	assert.doesNotThrow(() => verifyManifest({ npm: meta.npm, dir: meta.dir, version }));
});

const temp = mkdtempSync(join(tmpdir(), 'publish-released-'));
after(() => rmSync(temp, { recursive: true, force: true }));
const ui = { npm: '@x/ui', dir: 'packages/ui', version: '1.0.0' };
const versions = new Map([['@x/core', '0.4.2']]);
const writeUi = (dependencies, extra = {}) => {
	mkdirSync(join(temp, ui.dir), { recursive: true });
	writeFileSync(
		join(temp, ui.dir, 'package.json'),
		JSON.stringify({ name: ui.npm, version: ui.version, dependencies, ...extra }),
	);
};

test('a workspace range on a sibling is published as a caret range on its version', () => {
	writeUi(
		{ '@x/core': 'workspace:*', lit: '^3.0.0' },
		{ devDependencies: { '@x/core': 'workspace:^' } },
	);
	const manifest = publishManifest(ui, versions, temp);
	assert.deepEqual(manifest.dependencies, { '@x/core': '^0.4.2', lit: '^3.0.0' });
	assert.deepEqual(manifest.devDependencies, { '@x/core': '^0.4.2' });
});

test('a `*` range on a sibling (the UI on the core) is published as a caret range', () => {
	writeUi({ '@x/core': '*', lit: '^3.0.0' });
	assert.deepEqual(publishManifest(ui, versions, temp).dependencies, {
		'@x/core': '^0.4.2',
		lit: '^3.0.0',
	});
});

test('a workspace range on anything but a sibling cannot be published', () => {
	writeUi({ '@x/other': 'workspace:*' });
	assert.throws(() => verifyManifest(ui, versions, temp), /cannot be installed/);
	writeUi({ '@x/other': 'file:../other' });
	assert.throws(() => verifyManifest(ui, versions, temp), /cannot be installed/);
});

test('a real sibling range must be satisfied by the sibling version', () => {
	writeUi({ '@x/core': '^0.4.0' });
	assert.doesNotThrow(() => verifyManifest(ui, versions, temp));
	writeUi({ '@x/core': '^0.3.0' });
	assert.throws(() => verifyManifest(ui, versions, temp), /but that package is at 0.4.2/);
});
