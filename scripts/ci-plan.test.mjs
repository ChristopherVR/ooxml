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
	plan,
	shardsFor,
	viewersOfArea,
} from './ci-plan.mjs';

const viewers = (result) => result.viewers.map((viewer) => viewer.name);
const ALL_VIEWERS = Object.keys(VIEWERS);
/** The viewers checked in the shared `viewers` job; pptx has its own jobs (`plan.pptx`). */
const MATRIX_VIEWERS = ALL_VIEWERS.filter((key) => !VIEWERS[key].ownJobs);

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
		assert.ok(viewer.verify.length > 0, viewer.name);
		for (const command of viewer.verify) {
			const script = /^bun run ([w:.-]+)/u.exec(command)?.[1];
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

test('an engine change runs every pptx unit leg and the reference project; its tests run nothing', () => {
	const result = pptx('src/core/pptx/core/core/runtime/PptxHandlerRuntimeSmartArt.ts');
	assert.equal(result.tests.length, 11);
	assert.deepEqual(result.e2e.projects, ['react']);
	assert.equal(pptx('src/core/pptx/core/core/runtime/smartart.test.ts').run, false);
	assert.equal(pptx('src/core/xlsx/model.ts').run, false);
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
