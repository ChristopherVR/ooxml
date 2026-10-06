import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { test } from 'node:test';

import { CONSUMERS, VIEWERS, consumersOfArea, plan, shardsFor, viewersOfArea } from './ci-plan.mjs';

const names = (result) => result.consumers.map((consumer) => consumer.name);
const viewers = (result) => result.viewers.map((viewer) => viewer.name);
const ALL_VIEWERS = Object.keys(VIEWERS);

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

test('a pptx-only change skips the strict typecheck and checks only the external pptx viewer', () => {
	const result = plan(['src/core/pptx/converter/index.ts'], { testFiles: 5 });
	assert.equal(result.typecheck.strict, false);
	assert.equal(result.typecheck.pptx, true);
	assert.deepEqual(names(result), ['pptx-viewer']);
	assert.deepEqual(result.viewers, []);
});

test('a shared area can break every viewer', () => {
	const result = plan(['src/core/xml/parse.ts'], { testFiles: 300 });
	assert.equal(result.consumers.length, Object.keys(CONSUMERS).length);
	assert.deepEqual(viewers(result), ALL_VIEWERS);
	assert.deepEqual(result.test.shards, [1, 2, 3, 4, 5]);
});

test('a ui change checks the ui package and every viewer, but not the unit suite', () => {
	const result = plan(['src/ui/src/menu/context-menu.ts']);
	assert.equal(result.ui, true);
	assert.equal(result.test.run, false);
	assert.equal(result.consumers.length, Object.keys(CONSUMERS).length);
	assert.deepEqual(viewers(result), ALL_VIEWERS);
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
		assert.deepEqual(viewers(result), ALL_VIEWERS, file);
	}
});

test('a manual or scheduled run is always full', () => {
	const result = plan([], { full: true });
	assert.equal(result.full, true);
	assert.equal(result.consumers.length, Object.keys(CONSUMERS).length);
	assert.deepEqual(viewers(result), ALL_VIEWERS);
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
	assert.deepEqual(viewersOfArea('pptx'), []);
	assert.deepEqual(viewersOfArea('collab'), ALL_VIEWERS);
	assert.deepEqual(consumersOfArea('xlsx'), []);
	assert.deepEqual(consumersOfArea('pptx'), ['pptx']);
	assert.equal(consumersOfArea('collab').length, Object.keys(CONSUMERS).length);
});
