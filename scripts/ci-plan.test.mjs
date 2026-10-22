import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { test } from 'node:test';

import {
	CONSUMERS,
	PPTX_BINDINGS,
	PPTX_E2E_SHARDS,
	VIEWER_BROWSER_SHARDS,
	VIEWERS,
	consumersOfArea,
	onBasePlan,
	plan,
	plannedFiles,
	shardsFor,
	viewersOfArea,
} from './ci-plan.mjs';

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
	assert.deepEqual(
		result.e2e.matrix.map((entry) => entry.total),
		Array(PPTX_E2E_SHARDS.vue).fill(PPTX_E2E_SHARDS.vue),
	);
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
	assert.deepEqual(
		result.e2e.matrix,
		PPTX_BINDINGS.map((project) => ({ project, shard: 1, total: 1 })),
	);
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
		assert.equal(result.e2e.matrix.length, 14 + 4 * 7, file);
		assert.equal(result.e2e.specs, '', file);
	}
});

const verifyOf = (name, ...files) =>
	plan(files, { testFiles: 1 })
		.viewers.find((viewer) => viewer.name === name)
		?.verify.split('\n') ?? [];

const browserOf = (name, ...files) =>
	plan(files, { testFiles: 1 })
		.viewerBrowser.filter((entry) => entry.name === name)
		.flatMap((entry) => entry.verify.split('\n'));

test('a demo change runs only its viewer typecheck and browser tests', () => {
	const result = plan(['demos/docx/demo-vanilla/main.ts'], { testFiles: 1 });
	assert.deepEqual(viewers(result), ['docx']);
	assert.deepEqual(verifyOf('docx', 'demos/docx/demo-vanilla/main.ts'), ['bun run typecheck']);
	assert.deepEqual(browserOf('docx', 'demos/docx/demo-vanilla/main.ts'), [
		'bun run test:browser -- --shard=1/2',
		'bun run test:browser -- --shard=2/2',
	]);
	assert.equal(result.checks.core, false);
	assert.equal(result.pptx.run, false);
});

test('changed spec files alone run only those specs; support code runs the whole suite', () => {
	assert.deepEqual(browserOf('docx', 'e2e/docx/editor.spec.ts'), [
		'bun run test:browser -- editor.spec.ts',
	]);
	assert.equal(browserOf('docx', 'e2e/docx/helpers.ts').length, 2);
	// visio's test:browser also tests the packages, so it is never narrowed to specs
	assert.ok(
		browserOf('visio', 'e2e/visio/viewer.spec.ts').some((c) => c.includes('native-blocks')),
	);
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

test('a full pptx run gives react more shards than the other bindings', () => {
	const { matrix } = pptx('e2e/pptx/global-setup.ts').e2e;
	for (const project of PPTX_BINDINGS) {
		const entries = matrix.filter((entry) => entry.project === project);
		assert.equal(entries.length, PPTX_E2E_SHARDS[project], project);
		assert.deepEqual(
			entries.map((entry) => [entry.shard, entry.total]),
			entries.map((_, i) => [i + 1, PPTX_E2E_SHARDS[project]]),
			project,
		);
	}
	assert.ok(PPTX_E2E_SHARDS.react > PPTX_E2E_SHARDS.vue);
});

test('viewer browser suites run in their own sharded entries, not in the viewers list', () => {
	const result = plan([], { full: true });
	for (const viewer of result.viewers)
		assert.doesNotMatch(viewer.verify, /test:browser|playwright/u, viewer.name);
	for (const name of Object.keys(VIEWER_BROWSER_SHARDS)) {
		const entries = result.viewerBrowser.filter((entry) => entry.name === name);
		assert.equal(entries.length, VIEWER_BROWSER_SHARDS[name], name);
		entries.forEach((entry, i) => {
			assert.equal(entry.shard, i + 1);
			assert.equal(entry.total, VIEWER_BROWSER_SHARDS[name]);
			assert.equal(entry.dir, `viewers/${name}`);
			if (entry.total > 1)
				assert.match(entry.verify, new RegExp(`--shard=${i + 1}/${entry.total}`, 'u'));
			else assert.doesNotMatch(entry.verify, /--shard/u);
			assert.doesNotMatch(entry.verify, /playwright install/u);
		});
	}
});

test('visio shards the main Playwright run and keeps the chained runs on shard one', () => {
	const entries = plan([], { full: true }).viewerBrowser.filter((entry) => entry.name === 'visio');
	assert.match(entries[0].verify, /build:packages/u);
	assert.match(entries[0].verify, /native-blocks/u);
	assert.match(entries[0].verify, /test-packages-browser/u);
	assert.doesNotMatch(entries[1].verify, /native-blocks|build:packages|test-packages-browser/u);
	assert.match(entries[1].verify, /playwright test --shard=2\/2/u);
});

test('changed specs narrow a viewer browser run to one unsharded entry', () => {
	const result = plan(['e2e/docx/editing.spec.ts']);
	assert.deepEqual(
		result.viewerBrowser.map(({ name, shard, total }) => ({ name, shard, total })),
		[{ name: 'docx', shard: 1, total: 1 }],
	);
	assert.match(result.viewerBrowser[0].verify, /test:browser -- editing\.spec\.ts$/mu);
});

test('the browser jobs cache Playwright browsers and ci-success waits on the new jobs', () => {
	const workflow = readFileSync(new URL('../.github/workflows/ci.yml', import.meta.url), 'utf8');
	const job = (name) =>
		workflow.match(new RegExp(`^  ${name}:\\r?\\n([\\s\\S]*?)(?=^  \\S)`, 'm'))?.[1] ?? '';
	for (const name of ['pptx-e2e', 'pptx-e2e-packaged', 'viewer-browser']) {
		assert.match(job(name), /path: ~\/\.cache\/ms-playwright/u, name);
		assert.match(job(name), /hashFiles\('bun\.lock'\)/u, name);
	}
	assert.match(job('ui-test'), /--shard=\$\{\{ matrix\.shard \}\}\/3/u);
	const needs = /^  ci-success:[\s\S]*?needs:\s*\[([^\]]*)\]/m.exec(workflow)?.[1] ?? '';
	for (const name of ['ui', 'ui-test', 'viewers', 'viewer-browser'])
		assert.match(needs, new RegExp(`\\b${name}\\b`, 'u'), name);
});
