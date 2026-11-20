import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { test } from 'node:test';

import { plan } from './ci-plan.mjs';
import { classifyLock, parseLock, refineManifests } from './ci-manifests.mjs';
import { changedPackages, workspaceClosures, workspacesReaching } from './ci-lockfile.mjs';

const pkg = (id, meta = {}) => [id, '', meta, 'sha512-x'];
/** A lock with a docx package (left-pad), a pptx demo (on a workspace library) and the root. */
const lock = ({ leftPad = '1.0.0', zed = '2.0.0', extra = {} } = {}) => ({
	lockfileVersion: 1,
	workspaces: {
		'': { name: 'root', devDependencies: { oxlint: '^1' } },
		'viewers/docx/packages/react': { name: 'docx-react', dependencies: { 'left-pad': '^1' } },
		'demos/pptx/demo-vue': { name: 'pptx-vue-demo', dependencies: { lib: 'workspace:*' } },
		'viewers/pptx/packages/shared': { name: 'lib', dependencies: { zed: '^2' } },
	},
	packages: {
		oxlint: pkg('oxlint@1.0.0'),
		'left-pad': pkg(`left-pad@${leftPad}`, { dependencies: { 'repeat-string': '^1' } }),
		'repeat-string': pkg('repeat-string@1.6.1'),
		lib: pkg('lib@workspace:viewers/pptx/packages/shared'),
		zed: pkg(`zed@${zed}`),
		...extra,
	},
});
const text = (value) => JSON.stringify(value, null, 2);

test('a workspace installs its dependencies, theirs, and those of workspace packages it uses', () => {
	const closures = workspaceClosures(lock());
	assert.deepEqual([...closures.get('viewers/docx/packages/react')].sort(), [
		'left-pad',
		'repeat-string',
	]);
	assert.deepEqual([...closures.get('demos/pptx/demo-vue')].sort(), ['lib', 'zed']);
	assert.deepEqual([...closures.get('')], ['oxlint']);
});

test('a dependency nested under a workspace name resolves there before the root', () => {
	const nested = lock({ extra: { 'docx-react/left-pad': pkg('left-pad@0.9.0') } });
	assert.ok(
		workspaceClosures(nested).get('viewers/docx/packages/react').has('docx-react/left-pad'),
	);
});

test('a moved package reaches the workspaces whose trees hold it, before or after', () => {
	assert.deepEqual([...changedPackages(lock(), lock({ leftPad: '1.0.1' }))], ['left-pad']);
	assert.deepEqual(workspacesReaching(lock(), lock({ leftPad: '1.0.1' })), [
		'viewers/docx/packages/react',
	]);
	assert.deepEqual(workspacesReaching(lock(), lock({ zed: '2.1.0' })), [
		'demos/pptx/demo-vue',
		'viewers/pptx/packages/shared',
	]);
	// A package no workspace installs cannot be placed.
	assert.equal(
		workspacesReaching(lock(), lock({ extra: { orphan: pkg('orphan@1.0.0') } })),
		undefined,
	);
});

test('a dependency update plans as the manifests of the workspaces that install it', () => {
	const before = text(lock());
	const after = text(lock({ leftPad: '1.0.1' }));
	assert.deepEqual(classifyLock(before, after), {
		workspaces: ['viewers/docx/packages/react'],
		resolved: true,
	});
	const read = (path, side) =>
		path === 'bun.lock' ? (side === 'base' ? before : after) : undefined;
	const refined = refineManifests(['bun.lock'], read);
	assert.deepEqual(refined, { files: ['viewers/docx/packages/react/package.json'], deps: true });
	const result = plan(refined.files);
	assert.equal(result.full, false);
	assert.deepEqual(
		result.viewers.map((viewer) => viewer.name),
		['docx'],
	);
	assert.equal(result.pptx.run, false);
});

test('an update the root or the core installs still runs everything', () => {
	const root = lock();
	root.packages.oxlint = pkg('oxlint@1.1.0');
	assert.deepEqual(classifyLock(text(lock()), text(root)), { everything: true });
	const overrides = { ...lock(), overrides: { zed: '2.0.0' } };
	assert.deepEqual(classifyLock(text(lock()), text(overrides)), { everything: true });
});

test("every workspace of this repository's lock installs something the lock resolves", () => {
	const real = parseLock(readFileSync(new URL('../bun.lock', import.meta.url), 'utf8'));
	const closures = workspaceClosures(real);
	assert.equal(closures.size, Object.keys(real.workspaces).length);
	for (const [path, closure] of closures)
		if (Object.keys(real.workspaces[path]).some((key) => key.endsWith('ependencies')))
			assert.ok(closure.size > 0, `${path || 'the root'} resolves its dependencies`);
	assert.ok(closures.get('src/core').has('typescript'), 'the core installs its TypeScript');
});

test('a browser update the accessibility gates install also runs the byte budgets', () => {
	// perf-budget-routes.mjs loads the suite screens in e2e/a11y's Playwright.
	const result = plan(['e2e/a11y/package.json']);
	assert.equal(result.budgets, true);
	assert.equal(result.a11y.run, true);
});
