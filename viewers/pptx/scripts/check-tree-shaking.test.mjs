import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { after, test } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

import {
	droppedImportTimeEffects,
	findNamespaceCopies,
	isCliEntry,
	packageSideEffects,
	retainedBytes,
} from './check-tree-shaking.mjs';

const scriptsDir = dirname(fileURLToPath(import.meta.url));
const script = join(scriptsDir, 'check-tree-shaking.mjs');

const dirs = [];
after(() => {
	for (const dir of dirs) {
		rmSync(dir, { recursive: true, force: true });
	}
});

/** A throwaway `dist` holding `files` (name to source). */
function dist(files) {
	const dir = mkdtempSync(join(tmpdir(), 'tree-shaking-'));
	dirs.push(dir);
	for (const [name, source] of Object.entries(files)) {
		writeFileSync(join(dir, name), source);
	}
	return dir;
}

const editor = `function Editor() { return ${JSON.stringify('x'.repeat(5000))}; }`;
const canvas = 'export function Canvas() { return "canvas"; }';

test('an unannotated module-level component factory keeps the component', async () => {
	const dir = dist({
		'index.mjs': `import { forwardRef } from 'react';\n${editor}\nexport const Viewer = forwardRef(Editor);\n${canvas}\n`,
	});
	const bytes = await retainedBytes(dir, `export { Canvas } from '${dir}/index.mjs';`);
	assert.ok(bytes > 5000, `expected the editor to be kept, got ${bytes} bytes`);
});

test('a pure-annotated factory and display name let the component drop out', async () => {
	const dir = dist({
		'index.mjs': [
			"import { forwardRef } from 'react';",
			'function withDisplayName(c, n) { c.displayName = n; return c; }',
			editor,
			"export const Viewer = /* @__PURE__ */ withDisplayName(/* @__PURE__ */ forwardRef(Editor), 'Viewer');",
			canvas,
		].join('\n'),
	});
	const bytes = await retainedBytes(dir, `export { Canvas } from '${dir}/index.mjs';`);
	assert.ok(bytes < 200, `expected only the canvas to remain, got ${bytes} bytes`);
});

test('a display-name assignment statement keeps the component', async () => {
	const dir = dist({
		'index.mjs': `import { forwardRef } from 'react';\n${editor}\nexport const Viewer = /* @__PURE__ */ forwardRef(Editor);\nViewer.displayName = 'Viewer';\n${canvas}\n`,
	});
	const bytes = await retainedBytes(dir, `export { Canvas } from '${dir}/index.mjs';`);
	assert.ok(bytes > 5000, `expected the editor to be kept, got ${bytes} bytes`);
});

test('lazy chunks behind import() are not counted', async () => {
	const dir = dist({
		'index.mjs': `${canvas}\nexport const load = () => import('./lazy.mjs');\n`,
		'lazy.mjs': `export const big = ${JSON.stringify('y'.repeat(5000))};\n`,
	});
	const bytes = await retainedBytes(dir, `export { Canvas, load } from '${dir}/index.mjs';`);
	assert.ok(bytes < 500, `expected the lazy chunk to be excluded, got ${bytes} bytes`);
});

test('runtime namespace copies are reported', () => {
	const dir = dist({
		'chunk-a.mjs': "import * as star from 'pkg';\nvar ns = {};\n__reExport(ns, star);\n",
		'chunk-b.mjs': "export { a } from 'pkg';\n",
	});
	assert.deepEqual(findNamespaceCopies(dir), [join(dir, 'chunk-a.mjs')]);
});

/** A directory holding a symlink to this scripts folder, like `/tmp` on macOS. */
function linkedScriptsDir() {
	const dir = mkdtempSync(join(tmpdir(), 'tree-shaking-link-'));
	dirs.push(dir);
	const link = join(dir, 'scripts');
	symlinkSync(scriptsDir, link, 'dir');
	return link;
}

test('the CLI entry check sees through symlinked invocation paths', () => {
	const url = pathToFileURL(script).href;
	assert.equal(isCliEntry(url, script), true);
	assert.equal(isCliEntry(url, join(linkedScriptsDir(), 'check-tree-shaking.mjs')), true);
	assert.equal(isCliEntry(url, join(scriptsDir, 'check-viewer-props.mjs')), false);
	assert.equal(isCliEntry(url, undefined), false);
});

test('the check runs and fails when invoked through a symlinked path', () => {
	const dir = dist({
		'index.mjs': `${canvas}\nexport function SlideCanvas() {}\nexport function useViewerBuildingBlocks() {}\n`,
		'chunk-a.mjs': "import * as star from 'pkg';\nvar ns = {};\n__reExport(ns, star);\n",
	});
	const result = spawnSync(
		process.execPath,
		[join(linkedScriptsDir(), 'check-tree-shaking.mjs'), '--json', dir],
		{ encoding: 'utf8' },
	);
	assert.equal(result.status, 1, `expected exit 1, got ${result.status}: ${result.stderr}`);
	assert.match(result.stdout, /read-only canvas/);
	assert.match(result.stderr, /__reExport/);
});

/** A throwaway package with `manifest` as its package.json and `files` in its `dist`. */
function packageDist(manifest, files) {
	const root = mkdtempSync(join(tmpdir(), 'tree-shaking-package-'));
	dirs.push(root);
	writeFileSync(join(root, 'package.json'), JSON.stringify({ name: 'binding', ...manifest }));
	const dir = join(root, 'dist');
	mkdirSync(dir);
	for (const [name, source] of Object.entries(files)) {
		writeFileSync(join(dir, name), source);
	}
	return dir;
}

const registeringDist = {
	'index.mjs': "export { Canvas } from './chunk-canvas.mjs';\n",
	'chunk-canvas.mjs': [
		"import { defineThreeViewElement } from 'ooxml-ui/pptx';",
		'defineThreeViewElement();',
		canvas,
	].join('\n'),
};

test('the sideEffects field is read the way host bundlers read it', () => {
	const css = packageSideEffects(dirname(packageDist({ sideEffects: ['*.css'] }, {})));
	const root = dirname(packageDist({ sideEffects: ['./dist/register.mjs', 'dist/**/*.css'] }, {}));
	const listed = packageSideEffects(root);
	assert.equal(css('/pkg/dist/pptx-viewer.css'), true);
	assert.equal(css('/pkg/dist/chunk-canvas.mjs'), false);
	assert.equal(listed(join(root, 'dist/register.mjs')), true);
	assert.equal(listed(join(root, 'dist/themes/dark.css')), true);
	assert.equal(listed(join(root, 'dist/chunk-canvas.mjs')), false);
	assert.equal(packageSideEffects(dirname(packageDist({}, {})))('/any.mjs'), true);
	assert.equal(
		packageSideEffects(dirname(packageDist({ sideEffects: false }, {})))('/a.mjs'),
		false,
	);
});

test('a bare import keeps the three-view registration when sideEffects allows it', async () => {
	assert.deepEqual(await droppedImportTimeEffects(packageDist({}, registeringDist)), []);
	assert.deepEqual(
		await droppedImportTimeEffects(
			packageDist({ sideEffects: ['*.css', '*.mjs'] }, registeringDist),
		),
		[],
	);
});

test('a sideEffects list that leaves out the registering module drops the registration', async () => {
	assert.deepEqual(
		await droppedImportTimeEffects(packageDist({ sideEffects: ['*.css'] }, registeringDist)),
		['the <pptx-three-view> registration'],
	);
});

test('a pure-annotated registration call is reported as dropped', async () => {
	const annotated = packageDist(
		{},
		{
			...registeringDist,
			'chunk-canvas.mjs': registeringDist['chunk-canvas.mjs'].replace(
				'defineThreeViewElement();',
				'/* @__PURE__ */ defineThreeViewElement();',
			),
		},
	);
	assert.deepEqual(await droppedImportTimeEffects(annotated), [
		'the <pptx-three-view> registration',
	]);
});
