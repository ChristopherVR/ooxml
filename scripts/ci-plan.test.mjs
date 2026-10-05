import assert from 'node:assert/strict';
import { test } from 'node:test';

import { CONSUMERS, consumersOfArea, plan, shardsFor } from './ci-plan.mjs';

const names = (result) => result.consumers.map((consumer) => consumer.name);

test('a docs-only change runs nothing', () => {
	const result = plan(['docs/releasing.md', 'AGENTS.md']);
	assert.equal(result.full, false);
	assert.equal(result.test.run, false);
	assert.deepEqual(result.test.shards, []);
	assert.equal(result.typecheck.strict, false);
	assert.equal(result.build, false);
	assert.deepEqual(result.consumers, []);
});

test('a docx change tests the changed files, typechecks and checks only docx-viewer', () => {
	const result = plan(['src/docx/block-parser.ts'], { testFiles: 20 });
	assert.equal(result.test.mode, 'changed');
	assert.deepEqual(result.test.shards, [1]);
	assert.equal(result.typecheck.strict, true);
	assert.deepEqual(names(result), ['docx-viewer']);
});

test('a pptx-only change skips the strict typecheck', () => {
	const result = plan(['src/pptx/converter/index.ts'], { testFiles: 5 });
	assert.equal(result.typecheck.strict, false);
	assert.equal(result.typecheck.pptx, true);
	assert.deepEqual(names(result), ['pptx-viewer']);
});

test('a shared area can break every viewer', () => {
	const result = plan(['src/xml/parse.ts'], { testFiles: 300 });
	assert.equal(result.consumers.length, Object.keys(CONSUMERS).length);
	assert.deepEqual(result.test.shards, [1, 2, 3, 4, 5]);
});

test('a ui change checks the ui package and every viewer, but not the unit suite', () => {
	const result = plan(['packages/ui/src/menu/context-menu.ts']);
	assert.equal(result.ui, true);
	assert.equal(result.test.run, false);
	assert.equal(result.consumers.length, Object.keys(CONSUMERS).length);
});

test('a scripts or workflow-support change runs the script checks only', () => {
	const result = plan(['scripts/release-plan.mjs']);
	assert.equal(result.scripts, true);
	assert.equal(result.test.run, false);
	assert.equal(result.ui, false);
});

test('dependency, compiler and CI configuration runs everything', () => {
	for (const file of [
		'package.json',
		'bun.lock',
		'tsconfig.json',
		'.github/workflows/ci.yml',
		'vitest.config.ts',
	]) {
		const result = plan([file]);
		assert.equal(result.full, true, file);
		assert.equal(result.test.mode, 'all', file);
		assert.equal(result.test.shards.length, 6, file);
	}
});

test('a manual or scheduled run is always full', () => {
	const result = plan([], { full: true });
	assert.equal(result.full, true);
	assert.equal(result.consumers.length, Object.keys(CONSUMERS).length);
});

test('shards grow with the number of test files, to a cap of six', () => {
	assert.deepEqual(shardsFor(0, false), []);
	assert.deepEqual(shardsFor(1, false), [1]);
	assert.deepEqual(shardsFor(61, false), [1, 2]);
	assert.equal(shardsFor(10_000, false).length, 6);
	assert.equal(shardsFor(0, true).length, 6);
});

test('only viewer-specific areas narrow the viewers to check', () => {
	assert.deepEqual(consumersOfArea('xlsx'), ['xlsx']);
	assert.equal(consumersOfArea('collab').length, Object.keys(CONSUMERS).length);
});
