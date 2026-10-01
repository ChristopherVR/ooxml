import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { after, test } from 'node:test';

import {
	applyPlan,
	bumpVersion,
	commitLevel,
	isPublishedFile,
	PACKAGES,
	planRelease,
	satisfies,
} from './release-plan.mjs';

const CORE = PACKAGES.core.npm;
const UI = PACKAGES.ui.npm;
const roots = [];
after(() => roots.forEach((r) => rmSync(r, { recursive: true, force: true })));

const json = (data) => `${JSON.stringify(data, null, '\t')}\n`;

/**
 * A throwaway repository shaped like this one: core at the root (0.1.0), office-ui at packages/ui
 * (0.1.0, `workspace:*` on core). `tagged` lists the packages whose baseline tag is set.
 */
function repo({ ui = true, tagged = ['core', 'ui'], uiRange = 'workspace:*' } = {}) {
	const root = mkdtempSync(join(tmpdir(), 'release-plan-'));
	roots.push(root);
	const git = (...args) => execFileSync('git', args, { cwd: root, encoding: 'utf8' }).trim();
	const write = (file, text) => {
		mkdirSync(dirname(join(root, file)), { recursive: true });
		writeFileSync(join(root, file), text);
	};
	let n = 0;
	const commit = (message, files = {}) => {
		for (const [file, text] of Object.entries(files)) write(file, text);
		git('add', '-A');
		git('commit', '-q', '-m', message);
	};
	const touch = (file) => ({ [file]: `export const n = ${n++};\n` });
	git('init', '-q', '-b', 'main');
	for (const [k, v] of [
		['user.email', 't@example.com'],
		['user.name', 'T'],
		['commit.gpgsign', 'false'],
		['tag.gpgsign', 'false'],
	]) {
		git('config', k, v);
	}
	const files = {
		'package.json': json({ name: CORE, version: '0.1.0', scripts: { build: 'x' } }),
		...touch('src/index.ts'),
	};
	if (ui) {
		files['packages/ui/package.json'] = json({
			name: UI,
			version: '0.1.0',
			dependencies: { [CORE]: uiRange },
		});
		Object.assign(files, touch('packages/ui/src/index.ts'));
	}
	commit('feat: initial', files);
	for (const key of tagged) git('tag', `${PACKAGES[key].npm}@0.1.0`);
	const plan = (npm, npmHead) => planRelease({ root, packages: PACKAGES, npm, npmHead });
	return { root, git, write, commit, touch, plan, head: () => git('rev-parse', 'HEAD') };
}

const registry = (versions) => (name) => versions[name] ?? null;
const bothPublished = registry({ [CORE]: '0.1.0', [UI]: '0.1.0' });
const released = (p) => p.order.filter((k) => p.packages[k].release);

test('core is ordered before ui and a tagged HEAD releases nothing', () => {
	const r = repo();
	const p = r.plan(bothPublished);
	assert.deepEqual(p.order, ['core', 'ui']);
	assert.equal(p.anyChanged, false);
	assert.equal(p.packages.ui.dir, 'packages/ui');
	assert.equal(p.packages.ui.changelog, 'packages/ui/CHANGELOG.md');
	assert.equal(p.packages.core.changelog, 'CHANGELOG.md');
	assert.deepEqual(p.packages.ui.dependsOn, ['core']);
});

test('core-only change releases core as a patch and leaves ui alone', () => {
	const r = repo();
	r.commit('fix(core): handle empty input', r.touch('src/index.ts'));
	const p = r.plan(bothPublished);
	assert.deepEqual(released(p), ['core']);
	assert.equal(p.packages.core.version, '0.1.1');
	assert.equal(p.packages.ui.release, false);
});

test('a core minor at 0.x leaves the ui caret range, so ui is re-released (patch)', () => {
	const r = repo();
	r.commit('feat(core): add a parser', r.touch('src/index.ts'));
	const p = r.plan(bothPublished);
	assert.deepEqual(released(p), ['core', 'ui']);
	assert.equal(p.packages.core.version, '0.2.0');
	assert.equal(p.packages.ui.bump, 'patch', 'core commits never raise the ui bump level');
	assert.equal(p.packages.ui.version, '0.1.1');
	assert.match(p.packages.ui.reason, /dependency range/);
});

test('a core minor at >=1.0 stays inside ^1.x, so ui is not forced to release', () => {
	const r = repo();
	r.commit('chore: 1.0', {
		'package.json': json({ name: CORE, version: '1.0.0' }),
		'packages/ui/package.json': json({
			name: UI,
			version: '1.0.0',
			dependencies: { [CORE]: 'workspace:*' },
		}),
	});
	r.git('tag', '-f', `${CORE}@1.0.0`);
	r.git('tag', '-f', `${UI}@1.0.0`);
	r.commit('feat(core): add a parser', r.touch('src/index.ts'));
	const p = r.plan(registry({ [CORE]: '1.0.0', [UI]: '1.0.0' }));
	assert.deepEqual(released(p), ['core']);
	assert.equal(p.packages.core.version, '1.1.0');
});

test('a core major forces a ui release', () => {
	const r = repo();
	r.commit('chore: 1.0', {
		'package.json': json({ name: CORE, version: '1.0.0' }),
		'packages/ui/package.json': json({
			name: UI,
			version: '1.0.0',
			dependencies: { [CORE]: 'workspace:*' },
		}),
	});
	r.git('tag', '-f', `${CORE}@1.0.0`);
	r.git('tag', '-f', `${UI}@1.0.0`);
	r.commit('feat(core)!: drop the legacy entry', r.touch('src/index.ts'));
	const p = r.plan(registry({ [CORE]: '1.0.0', [UI]: '1.0.0' }));
	assert.deepEqual(released(p), ['core', 'ui']);
	assert.equal(p.packages.core.version, '2.0.0');
	assert.equal(p.packages.ui.version, '1.0.1');
});

test('ui-only change releases ui at its own level and leaves core alone', () => {
	const r = repo();
	r.commit('feat(ui): add a toolbar', r.touch('packages/ui/src/index.ts'));
	const p = r.plan(bothPublished);
	assert.deepEqual(released(p), ['ui']);
	assert.equal(p.packages.ui.bump, 'minor');
	assert.equal(p.packages.ui.version, '0.2.0');
	assert.equal(p.packages.ui.reason, 'own files changed');
});

test('both packages changed release together with independent levels', () => {
	const r = repo();
	r.commit('fix(core): a bug', r.touch('src/index.ts'));
	r.commit('feat(ui): a button', r.touch('packages/ui/src/index.ts'));
	const p = r.plan(bothPublished);
	assert.deepEqual(released(p), ['core', 'ui']);
	assert.equal(p.packages.core.version, '0.1.1');
	assert.equal(p.packages.ui.version, '0.2.0');
});

test('an untagged ui that is not on npm releases as initial at its manifest version', () => {
	const r = repo({ tagged: ['core'] });
	const p = r.plan(registry({ [CORE]: '0.1.0' }));
	assert.deepEqual(released(p), ['ui']);
	assert.equal(p.packages.ui.initial, true);
	assert.equal(p.packages.ui.bump, 'initial');
	assert.equal(p.packages.ui.version, '0.1.0');
	assert.equal(p.packages.ui.tag, `${UI}@0.1.0`);
	assert.equal(p.packages.ui.adopt, null);
});

test('offline, an untagged ui is still planned as initial', () => {
	const r = repo({ tagged: ['core'] });
	const p = r.plan(undefined);
	assert.equal(p.packages.ui.initial, true);
	assert.equal(p.packages.core.release, false);
});

test('a hand-published, untagged ui is a baseline, not a spurious bump', () => {
	const r = repo({ tagged: ['core'] });
	const p = r.plan(bothPublished);
	assert.equal(p.anyChanged, false);
	assert.equal(p.packages.ui.version, '0.1.0');
	assert.deepEqual(p.packages.ui.adopt, {
		tag: `${UI}@0.1.0`,
		sha: r.head(),
		fromRegistryHead: false,
	});
});

test('a hand-published ui is baselined at the commit npm recorded, so later changes release', () => {
	const r = repo({ tagged: ['core'] });
	const published = r.head();
	r.commit('feat(ui): add a toolbar', r.touch('packages/ui/src/index.ts'));
	const p = r.plan(bothPublished, (name) => (name === UI ? published : null));
	assert.deepEqual(released(p), ['ui']);
	assert.equal(p.packages.ui.adopt.sha, published);
	assert.equal(p.packages.ui.adopt.fromRegistryHead, true);
	assert.equal(p.packages.ui.bump, 'minor');
	assert.equal(p.packages.ui.version, '0.2.0');
});

test('a registry gitHead that is not in this history falls back to HEAD', () => {
	const r = repo({ tagged: ['core'] });
	const p = r.plan(bothPublished, () => 'deadbeefdeadbeefdeadbeefdeadbeefdeadbeef');
	assert.equal(p.packages.ui.adopt.sha, r.head());
	assert.equal(p.anyChanged, false);
});

test('a package that is not merged yet is left out of the plan', () => {
	const r = repo({ ui: false });
	const p = r.plan(registry({ [CORE]: '0.1.0' }));
	assert.deepEqual(p.order, ['core']);
	assert.equal(p.anyChanged, false);
});

test('tests, changelogs, docs and workspace wiring release nothing', () => {
	const r = repo();
	r.commit('feat(ui): add tests', {
		'packages/ui/src/index.test.ts': 'test\n',
		'packages/ui/CHANGELOG.md': '# Changelog\n',
		'CHANGELOG.md': '# Changelog\n',
		'docs/releasing.md': 'x\n',
	});
	const manifest = JSON.parse(readFileSync(join(r.root, 'package.json'), 'utf8'));
	r.commit('build: declare the workspace', {
		'package.json': json({ ...manifest, workspaces: ['packages/*'], scripts: { y: 'z' } }),
	});
	const ui = JSON.parse(readFileSync(join(r.root, 'packages/ui/package.json'), 'utf8'));
	r.commit('chore(ui): reorder the manifest', {
		'packages/ui/package.json': json(Object.fromEntries(Object.entries(ui).reverse())),
	});
	assert.equal(r.plan(bothPublished).anyChanged, false);
});

test('a ui dependency edit other than the sibling range is a release trigger', () => {
	const r = repo();
	const ui = JSON.parse(readFileSync(join(r.root, 'packages/ui/package.json'), 'utf8'));
	r.commit('fix(ui): depend on lit', {
		'packages/ui/package.json': json({
			...ui,
			dependencies: { ...ui.dependencies, lit: '^3.0.0' },
		}),
	});
	assert.deepEqual(released(r.plan(bothPublished)), ['ui']);
});

test('applyPlan stamps versions and keeps workspace ranges for Bun', () => {
	const r = repo();
	r.commit('feat(core): add a parser', r.touch('src/index.ts'));
	const p = r.plan(bothPublished);
	applyPlan({ root: r.root, packages: PACKAGES }, p);
	const read = (file) => JSON.parse(readFileSync(join(r.root, file), 'utf8'));
	assert.equal(read('package.json').version, '0.2.0');
	assert.equal(read('packages/ui/package.json').version, '0.1.1');
	assert.deepEqual(read('packages/ui/package.json').dependencies, { [CORE]: 'workspace:*' });
});

test('applyPlan keeps a `*` sibling range (published as a caret range later)', () => {
	const r = repo({ uiRange: '*' });
	r.commit('feat(core): add a parser', r.touch('src/index.ts'));
	applyPlan({ root: r.root, packages: PACKAGES }, r.plan(bothPublished));
	const ui = JSON.parse(readFileSync(join(r.root, 'packages/ui/package.json'), 'utf8'));
	assert.deepEqual(ui.dependencies, { [CORE]: '*' });
});

test('applyPlan repoints a plain range only for a released dependent', () => {
	const r = repo({ uiRange: '^0.1.0' });
	r.commit('feat(core): add a parser', r.touch('src/index.ts'));
	const p = r.plan(bothPublished);
	assert.equal(p.packages.ui.release, true, 'the 0.x minor leaves ^0.1.0');
	applyPlan({ root: r.root, packages: PACKAGES }, p);
	const ui = JSON.parse(readFileSync(join(r.root, 'packages/ui/package.json'), 'utf8'));
	assert.deepEqual(ui.dependencies, { [CORE]: '^0.2.0' });
});

test('the release commit does not retrigger once both packages are tagged', () => {
	const r = repo();
	r.commit('feat(core): add a parser', r.touch('src/index.ts'));
	const p = r.plan(bothPublished);
	applyPlan({ root: r.root, packages: PACKAGES }, p);
	r.commit('chore(release): bump versions and update changelogs [skip ci]', {
		'CHANGELOG.md': '# Changelog\n',
		'packages/ui/CHANGELOG.md': '# Changelog\n',
	});
	for (const key of ['core', 'ui']) r.git('tag', p.packages[key].tag);
	const next = r.plan(registry({ [CORE]: '0.2.0', [UI]: '0.1.1' }));
	assert.equal(next.anyChanged, false);
	assert.equal(next.packages.core.currentVersion, '0.2.0');
});

test('satisfies handles caret, tilde, exact and wildcard ranges', () => {
	assert.equal(satisfies('^0.1.0', '0.1.9'), true);
	assert.equal(satisfies('^0.1.0', '0.2.0'), false);
	assert.equal(satisfies('^1.2.0', '1.9.0'), true);
	assert.equal(satisfies('^1.2.0', '2.0.0'), false);
	assert.equal(satisfies('^1.2.0', '1.1.0'), false);
	assert.equal(satisfies('^0.0.3', '0.0.4'), false);
	assert.equal(satisfies('~1.2.0', '1.2.7'), true);
	assert.equal(satisfies('~1.2.0', '1.3.0'), false);
	assert.equal(satisfies('1.2.3', '1.2.3'), true);
	assert.equal(satisfies('*', '9.9.9'), true);
	assert.equal(satisfies('>=1.0.0', '1.0.0'), false);
});

test('helpers', () => {
	assert.equal(bumpVersion('1.2.3', 'minor'), '1.3.0');
	assert.equal(commitLevel('fix: x', 'BREAKING CHANGE: y'), 'major');
	assert.equal(commitLevel('feat(core): x'), 'minor');
	assert.equal(commitLevel('docs: x'), 'patch');
	assert.equal(isPublishedFile('packages/ui/src/a.test.ts'), false);
	assert.equal(isPublishedFile('packages/ui/src/a.ts'), true);
});
