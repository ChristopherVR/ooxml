import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { cpSync, mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import { test } from 'node:test';
import { VIEWER_PACKAGES, GLOBAL_TRIGGERS, planRelease } from './release-plan.mjs';
import { forbiddenManifestEntries } from './placeholder-manifest.mjs';

test('real release table publishes viewer and adapter changes but excludes docs and tests', () => {
	const root = mkdtempSync(resolve(tmpdir(), 'visio-release-table-'));
	const git = (...args) => execFileSync('git', args, { cwd: root, encoding: 'utf8' });
	git('init', '-q', '-b', 'main');
	git('config', 'user.name', 'Release test');
	git('config', 'user.email', 'test@example.com');
	git('config', 'commit.gpgsign', 'false');
	git('config', 'tag.gpgsign', 'false');
	for (const meta of Object.values(VIEWER_PACKAGES)) {
		mkdirSync(resolve(root, meta.dir), { recursive: true });
		cpSync(
			resolve(import.meta.dirname, '..', meta.dir, 'package.json'),
			resolve(root, meta.dir, 'package.json'),
		);
	}
	const change = (path, subject) => {
		mkdirSync(resolve(root, path, '..'), { recursive: true });
		writeFileSync(resolve(root, path), subject);
		git('add', '.');
		git('commit', '-q', '-m', subject);
	};
	change('src/controller.ts', 'feat(viewer): initialize');
	const baseline = () => {
		for (const meta of Object.values(VIEWER_PACKAGES)) git('tag', '-f', `${meta.npm}@0.0.2`);
	};
	baseline();
	const plan = () =>
		planRelease({
			root,
			packages: VIEWER_PACKAGES,
			globalTriggers: GLOBAL_TRIGGERS,
			npm: () => '0.0.2',
		});
	change('docs/index.html', 'docs: update guide');
	change('src/controller.test.ts', 'test: cover controller');
	assert.equal(plan().anyChanged, false);
	change('src/controller.ts', 'fix(viewer): update behavior');
	const viewerPlan = plan();
	for (const entry of Object.values(viewerPlan.packages)) {
		assert.equal(entry.manifest, `${entry.dir}/package.json`);
		assert.equal(entry.changelog, `${entry.dir}/CHANGELOG.md`);
	}
	assert.deepEqual(
		viewerPlan.order.filter((key) => viewerPlan.packages[key].release),
		['react', 'vue', 'angular', 'svelte', 'solid', 'vanilla'],
	);
	baseline();
	change('packages/bindings/src/common.ts', 'feat(bindings): forward options');
	const bindingPlan = plan();
	for (const key of ['react', 'vue', 'angular', 'svelte', 'solid', 'vanilla'])
		assert.equal(bindingPlan.packages[key].bump, 'minor');
	baseline();
	change('scripts/build-packages.mjs', 'fix(release): include workers');
	assert.ok(Object.values(plan().packages).every((entry) => entry.release));
});

test('published manifests allow registry peers and reject private packages', () => {
	assert.deepEqual(
		forbiddenManifestEntries({
			dependencies: { 'ooxml-core': '^0.10.0' },
			peerDependencies: { react: '>=18' },
		}),
		[],
	);
	assert.deepEqual(
		forbiddenManifestEntries({ dependencies: { '@christophervr/visio-viewer-bindings': '*' } }),
		['dependencies.@christophervr/visio-viewer-bindings'],
	);
});
