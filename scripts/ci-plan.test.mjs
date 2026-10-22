import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { test } from 'node:test';

import {
	CONSUMERS,
	PPTX_BINDINGS,
	PPTX_E2E_SHARDS,
	VIEWERS,
	consumersOfArea,
	onBasePlan,
	plan,
	plannedFiles,
	shardsFor,
	viewersOfArea,
} from './ci-plan.mjs';
import { pushBase } from './ci-base.mjs';
import {
	CORE_SHIPPING,
	classifyCoreManifest,
	classifyLock,
	parseLock,
	refineManifests,
} from './ci-manifests.mjs';

const viewers = (result) => result.viewers.map((viewer) => viewer.name);
const ALL_VIEWERS = Object.keys(VIEWERS);
/** The viewers checked in the shared `viewers` job; pptx has its own jobs (`plan.pptx`). */
const MATRIX_VIEWERS = ALL_VIEWERS.filter((key) => !VIEWERS[key].ownJobs);

test('UI typechecks reuse core declarations from the build job', () => {
	for (const file of [
		'src/ui/src/pptx/render/table-style.ts',
		'src/ui/src/teams/app/presentation-preview.ts',
		'src/ui/tsconfig.pptx.json',
		'.github/workflows/ci.yml',
	]) {
		const result = plan([file]);
		assert.equal(result.typecheck.ui, true);
		assert.equal(result.build, true);
	}
	const workflow = readFileSync(new URL('../.github/workflows/ci.yml', import.meta.url), 'utf8');
	const job = workflow.match(/^  typecheck-ui:\r?\n([\s\S]*?)(?=^  \S)/m)?.[1];
	assert.ok(job, 'the workflow has a UI typecheck job');
	assert.match(job, /needs:\s*\[[^\]]*\bbuild\b[^\]]*\]/u);
	assert.match(
		job,
		/download-artifact@\S+(?: +#[^\r\n]*)?\s+with:\s+name:\s*core-dist\s+path:\s*src\/core\/dist/u,
	);
	assert.ok(job.indexOf('download-artifact') < job.indexOf('src/ui typecheck'));
});

test('the viewers and the pptx jobs reuse ooxml-ui from the build job, not after the ui tests', () => {
	const workflow = readFileSync(new URL('../.github/workflows/ci.yml', import.meta.url), 'utf8');
	const job = (name) =>
		workflow.match(new RegExp(`^  ${name}:\\r?\\n([\\s\\S]*?)(?=^  \\S)`, 'm'))?.[1] ?? '';
	const build = job('build');
	assert.ok(build.indexOf('bun run build') < build.indexOf('bun run --cwd src/ui build'));
	assert.match(
		build,
		/upload-artifact@\S+(?: +#[^\r\n]*)?\s+with:\s+name:\s*ui-dist\s+path:\s*src\/ui\/dist/u,
	);
	for (const name of ['viewers', 'pptx-build']) {
		assert.match(job(name), /needs:\s*\[[^\]]*\bbuild\b[^\]]*\]/u, name);
		assert.doesNotMatch(job(name), /needs:\s*\[[^\]]*\bui\b|needs\.ui\./u, name);
	}
	// The ui job checks the build job's ooxml-ui instead of building its own.
	assert.doesNotMatch(job('ui'), /src\/ui build|upload-artifact/u);
});

test('a docs-only change runs nothing', () => {
	const result = plan(['docs/releasing.md', 'AGENTS.md']);
	assert.equal(result.full, false);
	assert.equal(result.test.run, false);
	assert.deepEqual(result.test.shards, []);
	assert.equal(result.typecheck.strict, false);
	assert.equal(result.build, false);
	assert.deepEqual(result.consumers, []);
	assert.deepEqual(result.viewers, []);
});

test('a docx change tests the changed files, typechecks and checks only the docx viewer', () => {
	const result = plan(['src/core/docx/block-parser.ts'], { testFiles: 20 });
	assert.equal(result.test.mode, 'changed');
	assert.deepEqual(result.test.shards, [1]);
	assert.equal(result.typecheck.strict, true);
	assert.deepEqual(viewers(result), ['docx']);
	assert.deepEqual(result.consumers, []);
});

test('a pptx-only change skips the strict typecheck and checks only the pptx viewer', () => {
	const result = plan(['src/core/pptx/converter/index.ts'], { testFiles: 5 });
	assert.equal(result.typecheck.strict, false);
	assert.equal(result.typecheck.pptx, true);
	assert.equal(result.pptx.run, true);
	assert.deepEqual(viewers(result), []);
	assert.deepEqual(result.consumers, []);
});

test('the pptx viewer runs in its own jobs, also for its demos and browser tests at the root', () => {
	for (const file of [
		'viewers/pptx/packages/react/src/index.ts',
		'demos/pptx/demo-vue/src/App.vue',
		'e2e/pptx/viewer-basics.spec.ts',
	]) {
		const result = plan([file], { testFiles: 0 });
		assert.equal(result.pptx.run, true, file);
		assert.deepEqual(viewers(result), [], file);
		assert.equal(result.build, true, file);
	}
	assert.equal(plan(['viewers/docx/src/index.ts'], { testFiles: 0 }).pptx.run, false);
});

test('a shared area can break every viewer', () => {
	const result = plan(['src/core/xml/parse.ts'], { testFiles: 300 });
	assert.equal(result.consumers.length, Object.keys(CONSUMERS).length);
	assert.deepEqual(viewers(result), MATRIX_VIEWERS);
	assert.equal(result.pptx.run, true);
	assert.deepEqual(result.test.shards, [1, 2, 3, 4, 5]);
});

test('a ui change checks the ui package and every viewer, but not the unit suite', () => {
	const result = plan(['src/ui/src/menu/context-menu.ts']);
	assert.equal(result.ui, true);
	assert.equal(result.test.run, false);
	assert.equal(result.consumers.length, Object.keys(CONSUMERS).length);
	assert.deepEqual(viewers(result), MATRIX_VIEWERS);
	assert.equal(result.pptx.run, true);
});

test('a change inside one viewer checks only that viewer, on a built core and ui', () => {
	const result = plan(['viewers/xlsx/packages/react/src/index.ts']);
	assert.deepEqual(viewers(result), ['xlsx']);
	assert.equal(result.viewers[0].dir, 'viewers/xlsx');
	assert.equal(result.test.run, false);
	assert.equal(result.scripts, false);
	assert.equal(result.build, true);
	assert.equal(result.ui, true);
	assert.deepEqual(result.consumers, []);
});

test('the root demos and browser tests of a viewer check that viewer', () => {
	assert.deepEqual(viewers(plan(['demos/xlsx/demo-vanilla/src.ts'])), ['xlsx']);
	assert.deepEqual(viewers(plan(['e2e/docx/editing.spec.ts'])), ['docx']);
});

test("a viewer's top-level notes and licence change nothing, its docs and scripts do", () => {
	assert.deepEqual(plan(['viewers/docx/README.md', 'viewers/docx/LICENSE']).viewers, []);
	assert.deepEqual(viewers(plan(['viewers/visio/docs/api.md'])), ['visio']);
	assert.deepEqual(viewers(plan(['viewers/teams/scripts/build-packages.mjs'])), ['teams']);
});

test('the shared package table is read by the visio scripts', () => {
	assert.deepEqual(viewers(plan(['scripts/viewer-packages.mjs'])), ['visio']);
});

test('every script a viewer verification runs exists in that viewer', () => {
	for (const viewer of Object.values(VIEWERS)) {
		const manifest = JSON.parse(readFileSync(join(viewer.dir, 'package.json'), 'utf8'));
		const commands = viewer.ownJobs
			? viewer.verify
			: [...Object.values(viewer.groups), ...Object.values(viewer.before ?? {})].flat();
		assert.ok(commands.length > 0, viewer.name);
		for (const command of commands) {
			const script = /^bun run ([\w:.-]+)/u.exec(command)?.[1];
			if (script) assert.ok(manifest.scripts?.[script], `${viewer.name}: ${command}`);
		}
	}
	for (const viewer of plan([], { full: true }).viewers)
		assert.equal(typeof viewer.verify, 'string');
});

test('a scripts or workflow-support change runs the script checks only', () => {
	const result = plan(['scripts/release-plan.mjs']);
	assert.equal(result.scripts, true);
	assert.equal(result.test.run, false);
	assert.equal(result.ui, false);
	assert.deepEqual(result.viewers, []);
});

test('dependency, compiler and CI configuration runs everything', () => {
	for (const file of [
		'package.json',
		'bun.lock',
		'src/core/package.json',
		'src/core/tsconfig.json',
		'.github/workflows/ci.yml',
		'src/core/vitest.config.ts',
	]) {
		const result = plan([file]);
		assert.equal(result.full, true, file);
		assert.equal(result.test.mode, 'all', file);
		assert.equal(result.test.shards.length, 6, file);
		assert.deepEqual(viewers(result), MATRIX_VIEWERS, file);
		assert.equal(result.pptx.run, true, file);
	}
});

test('a manual or scheduled run is always full', () => {
	const result = plan([], { full: true });
	assert.equal(result.full, true);
	assert.equal(result.consumers.length, Object.keys(CONSUMERS).length);
	assert.deepEqual(viewers(result), MATRIX_VIEWERS);
	assert.equal(result.pptx.run, true);
});

test('shards grow with the number of test files, to a cap of six', () => {
	assert.deepEqual(shardsFor(0, false), []);
	assert.deepEqual(shardsFor(1, false), [1]);
	assert.deepEqual(shardsFor(61, false), [1, 2]);
	assert.equal(shardsFor(10_000, false).length, 6);
	assert.equal(shardsFor(0, true).length, 6);
});

test('only format areas narrow the viewers to check', () => {
	assert.deepEqual(viewersOfArea('xlsx'), ['xlsx']);
	assert.deepEqual(viewersOfArea('pptx'), ['pptx']);
	assert.deepEqual(viewersOfArea('collab'), ALL_VIEWERS);
	assert.deepEqual(consumersOfArea('xlsx'), []);
	assert.deepEqual(consumersOfArea('pptx'), []);
	assert.equal(consumersOfArea('collab').length, Object.keys(CONSUMERS).length);
});

const pptx = (...files) => plan(files, { testFiles: 1 }).pptx;

test('a pptx binding change tests that package and runs only its own browser project', () => {
	const result = pptx('viewers/pptx/packages/vue/src/viewer/components/TableRenderer.vue');
	assert.deepEqual(result.tests, ['vue']);
	assert.deepEqual(result.e2e.projects, ['vue']);
	assert.equal(result.e2e.total, PPTX_E2E_SHARDS);
	assert.equal(result.e2e.specs, '');
	assert.equal(result.packaged, true);
});

test('a pptx test-only change runs its unit leg and no browser suite', () => {
	const result = pptx('viewers/pptx/packages/vue/src/viewer/table.test.ts');
	assert.deepEqual(result.tests, ['vue']);
	assert.deepEqual(result.e2e.projects, []);
	assert.equal(result.packaged, false);
});

test('shared pptx code tests every dependent package and runs the parity reference project', () => {
	const result = pptx('viewers/pptx/packages/shared/src/render/chart-axis.ts');
	assert.ok(result.tests.includes('shared') && result.tests.includes('svelte'));
	assert.ok(!result.tests.includes('core') && !result.tests.includes('tools'));
	assert.deepEqual(result.e2e.projects, ['react']);
});

test('the PowerPoint MCP server no longer drives shared UI or binding tests', () => {
	const result = pptx('viewers/pptx/packages/tools/src/server.ts');
	assert.deepEqual(result.tests, ['tools']);
	assert.deepEqual(result.e2e.projects, []);
	assert.equal(result.packaged, false);
});

test('an engine change runs every pptx unit leg and the reference project; its tests run nothing', () => {
	const result = pptx('src/core/pptx/core/core/runtime/PptxHandlerRuntimeSmartArt.ts');
	assert.equal(result.tests.length, 11);
	assert.deepEqual(result.e2e.projects, ['react']);
	assert.equal(pptx('src/core/pptx/core/core/runtime/smartart.test.ts').run, false);
	assert.equal(pptx('src/core/xlsx/model.ts').run, false);
});

test('another format area adds no pptx work, whatever else reaches the viewer', () => {
	// Each of these reaches pptx without running its engine legs; the docx and xlsx sources next to
	// them cannot break the viewer, so they must not turn the whole pptx run on.
	for (const reach of [
		'src/core/xml/lite.test.ts',
		'viewers/pptx/docs/guide/installation.md',
		'e2e/pptx/viewer-basics.spec.ts',
		'viewers/pptx/packages/cli/src/index.ts',
	]) {
		const result = pptx(reach, 'src/core/xlsx/model.ts', 'src/core/docx/model.ts');
		assert.deepEqual(result, pptx(reach), reach);
	}
	// A shared area next to them still reaches every leg, and so does a core file outside any area.
	assert.equal(pptx('src/core/xlsx/model.ts', 'src/core/xml/lite.ts').tests.length, 11);
	assert.equal(
		pptx('viewers/pptx/docs/guide/installation.md', 'src/core/index.ts').tests.length,
		11,
	);
});

test('changed spec files alone run just those specs, in every binding, on few shards', () => {
	const result = pptx('e2e/pptx/viewer-basics.spec.ts', 'e2e/pptx/smartart-reload.spec.ts');
	assert.deepEqual(result.tests, []);
	assert.deepEqual(result.e2e.projects, PPTX_BINDINGS);
	assert.equal(result.e2e.total, 1);
	assert.equal(result.e2e.specs, 'smartart-reload.spec.ts viewer-basics.spec.ts');
});

test('a demo runs its own binding; pptx docs run nothing', () => {
	assert.deepEqual(pptx('demos/pptx/demo-angular/src/app.component.ts').e2e.projects, ['angular']);
	assert.equal(pptx('viewers/pptx/docs/guide/installation.md').run, false);
});

test('suite-wide pptx files and CI configuration run the whole browser suite', () => {
	for (const file of [
		'e2e/pptx/global-setup.ts',
		'viewers/pptx/playwright.config.ts',
		'.github/workflows/ci.yml',
	]) {
		const result = pptx(file);
		assert.deepEqual(result.e2e.projects, PPTX_BINDINGS, file);
		assert.equal(result.e2e.total, PPTX_E2E_SHARDS, file);
		assert.equal(result.e2e.specs, '', file);
	}
});

const verifyOf = (name, ...files) =>
	plan(files, { testFiles: 1 })
		.viewers.find((viewer) => viewer.name === name)
		?.verify.split('\n') ?? [];

test('a demo change runs only its viewer typecheck and browser tests', () => {
	const result = plan(['demos/docx/demo-vanilla/main.ts'], { testFiles: 1 });
	assert.deepEqual(viewers(result), ['docx']);
	assert.deepEqual(verifyOf('docx', 'demos/docx/demo-vanilla/main.ts'), [
		'bun run typecheck',
		'bun x playwright install --with-deps chromium',
		'bun run test:browser',
	]);
	assert.equal(result.checks.core, false);
	assert.equal(result.pptx.run, false);
});

test('changed spec files alone run only those specs; support code runs the whole suite', () => {
	assert.ok(
		verifyOf('docx', 'e2e/docx/editor.spec.ts').includes('bun run test:browser -- editor.spec.ts'),
	);
	assert.ok(verifyOf('docx', 'e2e/docx/helpers.ts').includes('bun run test:browser'));
	// visio's test:browser also tests the packages, so it is never narrowed
	assert.ok(verifyOf('visio', 'e2e/visio/viewer.spec.ts').includes('bun run test:browser'));
});

test('a viewer test file, its MCP server and its docs run only what they can break', () => {
	assert.deepEqual(verifyOf('docx', 'viewers/docx/packages/react/src/a.test.ts'), [
		'bun run typecheck',
		'bun run test',
	]);
	const mcp = verifyOf('docx', 'viewers/docx/mcp/src/index.js');
	assert.equal(mcp.length, 1);
	assert.ok(mcp[0].includes('npm test'));
	assert.deepEqual(verifyOf('docx', 'viewers/docx/docs/guide.md'), []);
	assert.deepEqual(verifyOf('visio', 'viewers/visio/docs/site.test.mjs'), ['bun run test:docs']);
});

test('a product editor in ooxml-ui checks its own viewer; shared ui checks every viewer', () => {
	const editor = plan(['src/ui/src/docx/model-adapter.ts'], { testFiles: 1 });
	assert.deepEqual(viewers(editor), ['docx']);
	assert.equal(editor.pptx.run, false);
	const ribbon = plan(['src/ui/src/ribbon/ribbon.ts'], { testFiles: 1 });
	assert.deepEqual(viewers(ribbon), MATRIX_VIEWERS);
	assert.equal(ribbon.pptx.run, true);
});

test('a core test change runs the core tests and its typecheck, no viewer and no build', () => {
	const result = plan(['src/core/docx/parse.test.ts'], { testFiles: 1 });
	assert.equal(result.test.run, true);
	assert.deepEqual(result.viewers, []);
	assert.equal(result.pptx.run, false);
	assert.equal(result.build, false);
	assert.equal(result.typecheck.pptx, false);
});

test('teams builds its packages before its typecheck and browser tests', () => {
	const verify = verifyOf('teams', 'demos/teams/react/main.tsx');
	assert.equal(verify[0], 'bun run build:packages');
	assert.ok(verify.indexOf('bun run typecheck') > 0);
});

test('PowerPoint product UI reaches its own bindings without unrelated viewers', () => {
	const result = plan(['src/ui/src/pptx/render/group-drill.ts'], { testFiles: 1 });
	assert.equal(result.pptx.run, true);
	assert.equal(result.build, true);
	assert.deepEqual(viewers(result), []);
	assert.ok(result.pptx.tests.includes('angular'));
	assert.ok(result.pptx.tests.includes('vanilla'));
});

test('manual runs are full unless they ask for the changes since a base', () => {
	const diff = (base) => (base === 'abc' ? ['src/core/xlsx/model.ts'] : undefined);
	assert.equal(plannedFiles('schedule', 'abc', 'changed', diff), undefined);
	assert.equal(plannedFiles('workflow_dispatch', 'abc', 'full', diff), undefined);
	assert.equal(plannedFiles('workflow_dispatch', '', 'changed', diff), undefined);
	assert.equal(plannedFiles('workflow_dispatch', undefined, undefined, diff), undefined);
	assert.deepEqual(plannedFiles('workflow_dispatch', 'abc', 'changed', diff), [
		'src/core/xlsx/model.ts',
	]);
	// A base git cannot diff (not in a fork's history) falls back to a full run.
	assert.equal(plannedFiles('workflow_dispatch', 'missing', 'changed', diff), undefined);
	assert.deepEqual(plannedFiles('push', 'abc', undefined, diff), ['src/core/xlsx/model.ts']);
});

test('a run on the base tests the core sources the branch changed, with vitest related', () => {
	const changed = ['src/core/xlsx/model.ts', 'viewers/docx/docs/index.md'];
	const result = onBasePlan(plan(changed), changed);
	assert.equal(result.test.mode, 'related');
	assert.deepEqual(result.test.related, ['xlsx/model.ts']);
	const viewerOnly = ['viewers/docx/packages/react/src/index.ts'];
	const none = onBasePlan(plan(viewerOnly), viewerOnly);
	assert.equal(none.test.run, false);
	assert.deepEqual(none.test.shards, []);
});

const lock = (workspaces, packages = '"left-pad": ["left-pad@1.0.0", "", {}, "sha"],') => `{
  "lockfileVersion": 1,
  "workspaces": {
    ${workspaces}
  },
  "packages": {
    ${packages}
  },
}`;
const ws = (path, deps) => `"${path}": { "name": "x", "dependencies": { ${deps} }, },`;
const reader = (before, after) => (path, side) => (side === 'base' ? before[path] : after[path]);

test('a lockfile with trailing commas parses, commas inside strings survive', () => {
	assert.deepEqual(parseLock('{"a": ["x,}", 1,], "b": {"c": 2,},}'), {
		a: ['x,}', 1],
		b: { c: 2 },
	});
	assert.throws(() => parseLock('{"a": '));
});

test('a lock change confined to placeable workspaces plans as their manifests', () => {
	const before = lock(ws('demos/pptx/demo-host-review', '"a": "^1.0.0"'));
	const after = lock(ws('demos/pptx/demo-host-review', '"a": "^1.1.0"'));
	assert.deepEqual(classifyLock(before, after), { workspaces: ['demos/pptx/demo-host-review'] });
	const refined = refineManifests(
		['bun.lock', 'demos/pptx/demo-host-review/package.json'],
		reader({ 'bun.lock': before }, { 'bun.lock': after }),
	);
	assert.deepEqual(refined.files.sort(), ['demos/pptx/demo-host-review/package.json']);
	assert.equal(refined.deps, false);
	const result = plan(refined.files);
	assert.equal(result.full, false);
	assert.deepEqual(result.pptx.e2e.projects, []);
	assert.equal(plan(['bun.lock']).full, true);
});

test('a lock change to resolved packages, the root workspace or an unplaceable one runs everything', () => {
	const base = lock(ws('viewers/docx/packages/react', '"a": "^1"'));
	const resolved = lock(
		ws('viewers/docx/packages/react', '"a": "^1"'),
		'"left-pad": ["left-pad@1.0.1", "", {}, "sha2"],',
	);
	assert.deepEqual(classifyLock(base, resolved), { everything: true });
	assert.deepEqual(classifyLock(base, lock(ws('', '"a": "^1"'))), { everything: true });
	assert.deepEqual(classifyLock(base, lock(ws('apps/office-suite', '"a": "^1"'))), {
		everything: true,
	});
	assert.deepEqual(classifyLock(base, '{ not json'), { everything: true });
	assert.deepEqual(classifyLock(undefined, base), { everything: true });
	const refined = refineManifests(
		['bun.lock'],
		reader({ 'bun.lock': base }, { 'bun.lock': resolved }),
	);
	assert.deepEqual(refined, { files: ['bun.lock'], deps: true });
});

test('a core manifest change is inert, a shipping change or everything', () => {
	const manifest = (extra) =>
		JSON.stringify({
			name: 'ooxml-core',
			version: '1.0.0',
			exports: { '.': './a.js' },
			dependencies: { a: '1' },
			scripts: { build: 'x' },
			...extra,
		});
	assert.equal(classifyCoreManifest(manifest({}), manifest({ version: '1.0.1' })).kind, 'inert');
	assert.equal(
		classifyCoreManifest(manifest({}), manifest({ exports: { '.': './a.js', './b': './b.js' } }))
			.kind,
		'shipping',
	);
	assert.deepEqual(classifyCoreManifest(manifest({}), manifest({ dependencies: { a: '2' } })), {
		kind: 'everything',
		deps: true,
	});
	assert.deepEqual(classifyCoreManifest(manifest({}), manifest({ scripts: { build: 'y' } })), {
		kind: 'everything',
		deps: false,
	});
	assert.equal(classifyCoreManifest(manifest({}), '{').kind, 'everything');
	const read = (a, b) => reader({ 'src/core/package.json': a }, { 'src/core/package.json': b });
	assert.deepEqual(
		refineManifests(['src/core/package.json'], read(manifest({}), manifest({ version: '2.0.0' })))
			.files,
		[],
	);
	const shipping = refineManifests(
		['src/core/package.json'],
		read(manifest({}), manifest({ files: ['dist'] })),
	);
	assert.deepEqual(shipping.files, [CORE_SHIPPING]);
	const planned = plan(shipping.files);
	assert.equal(planned.full, false);
	assert.equal(planned.build, true);
	assert.equal(planned.checks.core, true);
	assert.equal(planned.typecheck.strict, true);
	assert.equal(planned.test.run, true);
	assert.equal(planned.test.mode, 'changed');
	assert.deepEqual(viewers(planned), []);
	assert.equal(planned.pptx.run, false);
	assert.equal(plan(['src/core/package.json']).full, true);
});

test('a configuration full run uses the reference project, a dependency or nightly run all five', () => {
	const config = plan(['src/core/tsconfig.json'], { reference: true });
	assert.equal(config.full, true);
	assert.deepEqual(config.pptx.e2e.projects, ['react']);
	assert.equal(config.pptx.e2e.total, PPTX_E2E_SHARDS);
	assert.match(config.pptx.e2e.reason, /reference only/u);
	const deps = plan(['bun.lock']);
	assert.deepEqual(deps.pptx.e2e.projects, PPTX_BINDINGS);
	assert.match(deps.pptx.e2e.reason, /all bindings/u);
	const nightly = plan([], { full: true });
	assert.equal(nightly.pptx.e2e.projects.length, PPTX_BINDINGS.length);
	// A binding's own change on top of a config change still runs that binding.
	const both = plan(['src/core/tsconfig.json', 'demos/pptx/demo-vue/main.ts'], {
		reference: true,
	});
	assert.deepEqual(both.pptx.e2e.projects, ['react', 'vue']);
	assert.equal(plan(['viewers/pptx/packages/vue/src/a.ts']).pptx.e2e.reason, '');
});

const commits = ['head', 'c3', 'c2', 'c1', 'c0'];
const git =
	() =>
	(...args) => {
		if (args[0] === 'rev-parse') return 'head';
		if (args[0] === 'rev-list') return commits.join('\n');
		if (args[0] === 'merge-base') {
			const [, , candidate, ref] = args;
			if (commits.indexOf(candidate) >= commits.indexOf(ref)) return '';
			throw new Error('not an ancestor');
		}
		throw new Error(args.join(' '));
	};
const runs =
	(...shas) =>
	async () => ({
		workflow_runs: shas.map((sha) => ({ head_sha: sha, conclusion: 'success', event: 'push' })),
	});
const lookup = (before, fetchJson, token = 't') =>
	pushBase({ before, repo: 'o/r', token, exec: git(), fetchJson });

test('a push plans from the newest green ancestor when it is older than the event base', async () => {
	assert.deepEqual(await lookup('c3', runs('c1', 'c0')), { base: 'c1', source: 'green' });
	// The event's own base is the green commit, or older than it: keep it.
	assert.deepEqual(await lookup('c1', runs('c1')), { base: 'c1', source: 'before' });
	assert.deepEqual(await lookup('c1', runs('c3')), { base: 'c1', source: 'before' });
	// HEAD's own run is not a base, and nothing green means a full run.
	assert.deepEqual(await lookup('c3', runs('head')), { base: '', source: 'none' });
});

test('a failed lookup, a missing token or non-success runs fall back safely', async () => {
	const fail = async () => {
		throw new Error('503');
	};
	assert.deepEqual(await lookup('c3', fail), { base: 'c3', source: 'before' });
	assert.deepEqual(await lookup('c3', runs('c0'), ''), { base: 'c3', source: 'before' });
	const other = async () => ({
		workflow_runs: [
			{ head_sha: 'c0', conclusion: 'failure', event: 'push' },
			{ head_sha: 'c1', conclusion: 'success', event: 'pull_request' },
		],
	});
	assert.deepEqual(await lookup('c3', other), { base: '', source: 'none' });
});
