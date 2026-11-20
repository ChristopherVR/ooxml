// The planner's import-graph rules (graphReach in ci-plan.mjs): a shared core or UI change runs
// what imports it, never less, and falls back to the area rules when the graph has no answer.
import assert from 'node:assert/strict';
import { test } from 'node:test';

import { PPTX_BINDINGS, plan } from './ci-plan.mjs';

/** A graph stub: the owner keys each file reaches (ownerOf in ci-graph.mjs); others have no answer. */
const stub = (answers) => (file) => (file in answers ? new Set(answers[file]) : undefined);
const viewers = (result) => result.viewers.map((viewer) => viewer.name).sort();
const browsers = (result) => [...new Set(result.viewerBrowser.map((entry) => entry.name))].sort();

test('a shared core change runs only the viewers whose files import it', () => {
	const file = 'src/core/chart/edit/series.ts';
	const result = plan([file], { reach: stub({ [file]: ['xlsx', 'teams'] }) });
	assert.deepEqual(viewers(result), ['teams', 'xlsx']);
	assert.deepEqual(browsers(result), ['teams', 'xlsx']);
	assert.equal(result.pptx.run, false);
	// Without the graph the area rules reach every viewer.
	assert.deepEqual(viewers(plan([file])), ['docx', 'teams', 'visio', 'xlsx']);
	assert.equal(plan([file]).pptx.run, true);
});

test('a change only an MCP server imports runs that check, not the browser suite', () => {
	const file = 'src/core/automation/outline.ts';
	const result = plan([file], { reach: stub({ [file]: ['docx:mcp', 'xlsx:mcp'] }) });
	const docx = result.viewers.find((viewer) => viewer.name === 'docx');
	assert.match(docx.verify, /copy-mcp-sources/);
	assert.doesNotMatch(docx.verify, /typecheck|test:browser/);
	assert.deepEqual(browsers(result), []);
});

test('the graph adds a viewer the area rules miss: a format area another product embeds', () => {
	// The area rule says src/core/docx reaches only the Word viewer; Teams embeds the Word editor.
	const file = 'src/core/docx/document-parts.ts';
	assert.deepEqual(viewers(plan([file])), ['docx']);
	const result = plan([file], { reach: stub({ [file]: ['docx', 'teams', 'xlsx:mcp'] }) });
	assert.deepEqual(viewers(result), ['docx', 'teams', 'xlsx']);
	assert.deepEqual(browsers(result), ['docx', 'teams']);
});

test('PowerPoint runs the unit legs that import a change, and its browser suite only if the demos do', () => {
	const file = 'src/ui/src/pptx/render/table-style.ts';
	const legsOnly = plan([file], { reach: stub({ [file]: ['pptx:shared'] }) });
	assert.equal(legsOnly.pptx.run, true);
	// shared and every package that depends on it.
	assert.deepEqual(legsOnly.pptx.tests, [
		'shared',
		'locales',
		'react',
		'react18',
		'vue',
		'angular',
		'vanilla',
		'svelte',
		'cli',
	]);
	assert.deepEqual(legsOnly.pptx.e2e.projects, []);
	const demos = plan([file], { reach: stub({ [file]: ['pptx', 'pptx:vue'] }) });
	assert.deepEqual(demos.pptx.e2e.projects, ['react']);
	assert.deepEqual(demos.pptx.tests, ['vue']);
	assert.equal(demos.pptx.packaged, true);
});

test('a file no viewer imports, or one the graph cannot speak for, keeps the area rules', () => {
	const unused = 'src/core/pptx/core/ppt/writer/colors.ts';
	const result = plan([unused], { reach: stub({ [unused]: ['onedrive'] }) });
	assert.equal(result.pptx.run, true);
	assert.deepEqual(result.pptx.e2e.projects, ['react']);
	const data = 'src/core/text/font-metrics/advances.json';
	assert.deepEqual(viewers(plan([data], { reach: stub({}) })), viewers(plan([data])));
});

test('a full run ignores the graph and runs everything', () => {
	const file = 'src/core/chart/edit/series.ts';
	const result = plan([file], { full: true, reach: stub({ [file]: ['xlsx'] }) });
	assert.deepEqual(viewers(result), ['docx', 'teams', 'visio', 'xlsx']);
	assert.deepEqual(result.pptx.e2e.projects, PPTX_BINDINGS);
});
