import assert from 'node:assert/strict';
import { test } from 'node:test';

import { importGraph, ownerOf, repositoryGraph, specifiers, uiEntries } from './ci-graph.mjs';

/** A graph over an in-memory tree: `files` maps paths to text, `packages` names to folders. */
const graph = (files, packages = {}, entries = new Map()) =>
	importGraph(new Map(Object.entries(files)), new Map(Object.entries(packages)), entries);
const sorted = (set) => (set ? [...set].sort() : set);

test('reads static, dynamic, worker, require and stylesheet imports, but not commented examples', () => {
	const { exact, partial } = specifiers(`
// import fake from './commented-out';
/**
 * {@link import('./doc-link')}
 */
import a from './a';
import type { B } from "../b";
export * from './c';
import './side-effect';
const d = await import('./d');
const w = new Worker(new URL('./worker.ts', import.meta.url));
const r = require('./r');
const locale = await import(\`./locales/\${name}\`);
const url = 'https://example.org/x'; // a URL is not a comment
@import url('./theme.css');
`);
	assert.deepEqual(exact.sort(), [
		'../b',
		'./a',
		'./c',
		'./d',
		'./r',
		'./side-effect',
		'./theme.css',
		'./worker.ts',
	]);
	assert.deepEqual(partial, ['./locales/']);
});

test('names who runs a viewer path: the viewer, its MCP server or one PowerPoint package', () => {
	assert.equal(ownerOf('viewers/docx/packages/core/src/index.ts'), 'docx');
	assert.equal(ownerOf('viewers/docx/mcp/src/index.js'), 'docx:mcp');
	assert.equal(ownerOf('viewers/pptx/packages/react/src/index.ts'), 'pptx:react');
	assert.equal(ownerOf('viewers/pptx/playwright.config.ts'), 'pptx');
	assert.equal(ownerOf('demos/xlsx/main.ts'), 'xlsx');
	assert.equal(ownerOf('e2e/visio/smoke.spec.ts'), 'visio');
	assert.equal(ownerOf('src/core/xlsx/model.ts'), undefined);
});

test('a viewer reaches what its files import, through package subpaths and re-exports', () => {
	const reach = graph({
		'src/core/a/index.ts': "export * from './x';",
		'src/core/a/x.ts': 'export const x = 1;',
		'src/core/b/index.ts': 'export const b = 1;',
		'viewers/docx/packages/core/src/index.ts': "import { x } from 'ooxml-core/a';",
		'viewers/xlsx/packages/core/src/index.ts': "import { b } from 'ooxml-core/b';",
	});
	assert.deepEqual(sorted(reach('src/core/a/x.ts')), ['docx']);
	assert.deepEqual(sorted(reach('src/core/b/index.ts')), ['xlsx']);
});

test('an import of another workspace package reaches every file of that package', () => {
	const reach = graph(
		{
			'viewers/docx/packages/core/src/index.ts': "export * from './a';",
			'viewers/docx/packages/core/src/a.ts': "import 'ooxml-core/a';",
			'viewers/docx/packages/core/src/unused.ts': "import 'ooxml-core/b';",
			'src/core/a/index.ts': '',
			'src/core/b/index.ts': '',
			'demos/teams/main.ts': "import { Viewer } from 'docx-core';",
		},
		{ 'docx-core': 'viewers/docx/packages/core' },
	);
	assert.deepEqual(sorted(reach('src/core/b/index.ts')), ['docx', 'teams']);
});

test('an ooxml-ui subpath resolves through the tsup entries, or its source path', () => {
	const entries = uiEntries([
		[
			'src/ui/tsup.config.ts',
			"entry: { index: 'src/index.ts', 'chart-format-locales': 'src/chart/locales/index.ts' }",
		],
		['src/ui/tsup.pptx.config.ts', "entry: { 'dom/index': 'src/pptx/dom/index.ts' }"],
	]);
	assert.equal(entries.get('chart-format-locales'), 'src/ui/src/chart/locales/index.ts');
	assert.equal(entries.get('pptx/dom'), 'src/ui/src/pptx/dom/index.ts');
	assert.equal(entries.get(''), 'src/ui/src/index.ts');
	const reach = graph(
		{
			'src/ui/src/chart/locales/index.ts': '',
			'src/ui/src/pptx/dom/index.ts': '',
			'src/ui/src/theme.ts': '',
			'viewers/xlsx/packages/core/src/index.ts': "import 'ooxml-ui/chart-format-locales';",
			'viewers/pptx/packages/vanilla/src/index.ts':
				"import 'ooxml-ui/pptx/dom'; import 'ooxml-ui/theme';",
		},
		{},
		entries,
	);
	assert.deepEqual(sorted(reach('src/ui/src/chart/locales/index.ts')), ['xlsx']);
	assert.deepEqual(sorted(reach('src/ui/src/pptx/dom/index.ts')), ['pptx:vanilla']);
	assert.deepEqual(sorted(reach('src/ui/src/theme.ts')), ['pptx:vanilla']);
});

test('a specifier that resolves to no file reaches everything in the folder it points into', () => {
	const reach = graph({
		'src/core/b/one.ts': '',
		'src/core/b/two.ts': '',
		'src/core/c/three.ts': '',
		'viewers/visio/packages/core/src/index.ts': "import '../../../../../src/core/b/generated';",
		'viewers/docx/packages/core/src/locale.ts':
			'const m = await import(`../../../../../src/core/c/${name}`);',
		'viewers/xlsx/packages/core/src/index.ts': "import 'ooxml-core/not-an-area';",
	});
	assert.deepEqual(sorted(reach('src/core/b/two.ts')), ['visio', 'xlsx']);
	assert.deepEqual(sorted(reach('src/core/c/three.ts')), ['docx', 'xlsx']);
});

test("a folder guess stays among its owner's files, and MCP servers reach as their own key", () => {
	const reach = graph({
		'viewers/xlsx/vite.config.ts': "const root = new URL('.', import.meta.url);",
		'viewers/xlsx/mcp/src/index.js': "import 'ooxml-core/automation';",
		'viewers/docx/mcp/src/index.js': "import 'ooxml-core/automation';",
		'src/core/automation/index.ts': '',
	});
	assert.deepEqual(sorted(reach('src/core/automation/index.ts')), ['docx:mcp', 'xlsx:mcp']);
});

test('package tests, build scripts and viewer maintenance scripts lead nowhere on their own', () => {
	const reach = graph({
		'src/core/a/index.ts': "export const a = 1; // see './a.test'",
		'src/core/a/a.test.ts': "import '../b/index';",
		'src/core/b/index.ts': '',
		'viewers/pptx/scripts/measure.mjs': "import 'ooxml-core/b';",
		'viewers/docx/packages/core/src/index.ts': "import '../../../../../src/core/a/missing';",
	});
	// The folder guess into src/core/a does not walk through the test, and the script is never run.
	assert.deepEqual(sorted(reach('src/core/b/index.ts')), []);
	assert.equal(reach('src/core/scripts/build.mjs'), undefined);
	assert.equal(reach('src/core/a/data.json'), undefined);
});

test('on this repository, each viewer reaches its own core area and the shared areas', () => {
	const reach = repositoryGraph();
	for (const viewer of ['docx', 'xlsx', 'visio', 'teams'])
		assert.ok(reach(`src/core/${viewer}/index.ts`).has(viewer), `${viewer} reaches its area`);
	assert.ok(reach('src/core/pptx/index.ts').has('pptx'));
	for (const viewer of ['docx', 'xlsx', 'visio', 'teams', 'pptx'])
		assert.ok(
			reach('src/core/units/index.ts').has(viewer) ||
				reach('src/core/units/index.ts').has(`${viewer}:mcp`),
			`${viewer} reaches units`,
		);
	// The Teams app is Teams' own: no other viewer embeds it.
	assert.deepEqual(
		[...reach('src/ui/src/teams/app/index.ts')].filter((key) => !key.startsWith('teams')),
		[],
	);
});
