import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { readdir, readFile } from 'node:fs/promises';
import { test } from 'node:test';

import { PACKAGES } from './release-plan.mjs';
import {
	provenanceEnabled,
	publishArgs,
	resolveTargets,
	verifyManifest,
} from './publish-released.mjs';
import { workspacePackageNames } from './workspace-packages.mjs';

const [key, meta] = Object.entries(PACKAGES)[0];
const version = JSON.parse(await readFile(`${meta.dir}/package.json`, 'utf8')).version;

test('a tag resolves to exactly one package', () => {
	const [target] = resolveTargets({ tag: `${meta.npm}@1.2.3` });
	assert.deepEqual(target, { key, npm: meta.npm, dir: meta.dir, version: '1.2.3' });
	const [server] = resolveTargets({ tag: 'openteams-server@0.2.0' });
	assert.deepEqual(server, {
		key: 'server',
		npm: 'openteams-server',
		dir: 'server',
		version: '0.2.0',
	});
});

test('a malformed or unknown tag is rejected', () => {
	assert.throws(() => resolveTargets({ tag: 'v1.2.3' }), /Invalid tag|Unknown package/);
	assert.throws(() => resolveTargets({ tag: 'teams-viewer@1.0.0' }), /Unknown package/);
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

test('provenance only inside GitHub Actions with OIDC and when asked for', () => {
	const ci = { GITHUB_ACTIONS: 'true', ACTIONS_ID_TOKEN_REQUEST_URL: 'https://example.invalid' };
	assert.equal(provenanceEnabled({}), false);
	assert.equal(provenanceEnabled({ NPM_PROVENANCE: 'true' }), false, 'never on a laptop');
	assert.equal(provenanceEnabled({ ...ci }), false, 'off unless asked (private repository)');
	assert.equal(provenanceEnabled({ ...ci, NPM_PROVENANCE: 'false' }), false);
	assert.equal(provenanceEnabled({ ...ci, NPM_PROVENANCE: 'true' }), true);
	assert.ok(publishArgs({ tag: 'latest', provenance: true }).includes('--provenance'));
	assert.ok(!publishArgs({ tag: 'latest' }).includes('--provenance'));
	assert.deepEqual(publishArgs({ tag: 'old', dryRun: true }).slice(0, 5), [
		'publish',
		'--access',
		'public',
		'--tag',
		'old',
	]);
	assert.ok(publishArgs({ tag: 'latest', dryRun: true }).includes('--dry-run'));
	assert.ok(!publishArgs({ tag: 'latest' }).some((arg) => /otp/iu.test(arg)));
});

test('exactly the packages in the release table are public; every other workspace package is private', async () => {
	const published = new Set(
		Object.values(PACKAGES)
			.filter((p) => p.dir.startsWith('packages/'))
			.map((p) => p.dir.replace('packages/', '')),
	);
	assert.equal(published.size, 6);
	const server = JSON.parse(await readFile('server/package.json', 'utf8'));
	assert.doesNotThrow(() =>
		verifyManifest({ npm: server.name, dir: 'server', version: server.version }),
	);
	// Leftovers of removed packages (untracked dist or node_modules) are skipped; a directory with
	// source but no manifest fails here.
	for (const dir of await workspacePackageNames(new URL('../packages/', import.meta.url))) {
		const manifest = JSON.parse(await readFile(`packages/${dir}/package.json`, 'utf8'));
		assert.equal(
			Boolean(manifest.private),
			!published.has(dir),
			`packages/${dir}: ${published.has(dir) ? 'must be public' : 'must be private'}`,
		);
		if (published.has(dir))
			assert.doesNotThrow(() =>
				verifyManifest({ npm: manifest.name, dir: `packages/${dir}`, version: manifest.version }),
			);
	}
	for (const dir of await readdir('demos', { withFileTypes: true })) {
		if (!dir.isDirectory() || !existsSync(`demos/${dir.name}/package.json`)) continue;
		const manifest = JSON.parse(await readFile(`demos/${dir.name}/package.json`, 'utf8'));
		assert.equal(manifest.private, true, `demos/${dir.name} must be private`);
	}
});
