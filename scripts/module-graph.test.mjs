import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';

import { walkModuleGraph } from './module-graph.mjs';

/** Write `files` into a fresh directory, walk from `entry.*`, then clean up. */
async function walk(files, options) {
	const directory = mkdtempSync(join(tmpdir(), 'ooxml-module-graph-'));
	try {
		for (const [name, source] of Object.entries(files)) {
			writeFileSync(join(directory, name), source);
		}
		const entry = Object.keys(files).find((name) => name.startsWith('entry.'));
		const graph = await walkModuleGraph(join(directory, entry), options);
		const local = (file) => file.slice(directory.length + 1);
		return {
			files: graph.files.map(local),
			lazy: graph.lazy.map(local),
			// esbuild parses files in parallel, so the order across files is not stable.
			bare: graph.bare
				.map(({ specifier, kind, importer }) => `${kind} ${specifier} <- ${local(importer)}`)
				.sort(),
		};
	} finally {
		rmSync(directory, { recursive: true, force: true });
	}
}

test('imports used only as types are erased, value imports are kept', async () => {
	const { bare } = await walk({
		'entry.ts': `import type { A } from 'types-only';
import { type B, type C } from 'inline-types';
export type { D } from 'reexported-types';
import { Unused } from 'used-as-type';
import { type E, value } from 'mixed';
import * as ns from 'namespace';
export { x } from 'reexport';
export * from 'star';
import 'side-effect';
export const run = (u: Unused, e: E): unknown => [value, ns, u, e];
`,
	});
	assert.deepEqual(bare, [
		'import-statement mixed <- entry.ts',
		'import-statement namespace <- entry.ts',
		'import-statement reexport <- entry.ts',
		'import-statement side-effect <- entry.ts',
		'import-statement star <- entry.ts',
	]);
});

test('dynamic imports and require calls count; type queries, comments and strings do not', async () => {
	const { bare } = await walk({
		'entry.js': `// import { commented } from 'line-comment';
/* import('block-comment') */
const glob = '/*';
const lazy = () => import('lazy');
const cjs = require('required');
const text = "import { inString } from 'string'";
export { glob, lazy, cjs, text };
`,
		'types.ts': '',
	});
	assert.deepEqual(bare, ['dynamic-import lazy <- entry.js', 'require-call required <- entry.js']);
	const typed = await walk({
		'entry.ts': `export type UseChat = typeof import('type-query').useChat;`,
	});
	assert.deepEqual(typed.bare, []);
});

test('the walk follows relative modules and can stop at lazy boundaries', async () => {
	const files = {
		'entry.ts': `import { value } from './helper';
import type { Heavy } from './types';
import styles from './styles.css?raw';
export const load = () => import('./panel');
export const use = (heavy: Heavy) => [value, styles, heavy];
`,
		'helper.ts': `export { value } from 'light';`,
		'types.ts': `import { schema } from 'heavy'; export type Heavy = typeof schema;`,
		'panel.ts': `export { runtime } from 'heavy';`,
	};
	const all = await walk(files);
	assert.deepEqual(all.bare, [
		'import-statement heavy <- panel.ts',
		'import-statement light <- helper.ts',
	]);
	assert.deepEqual(all.lazy, []);
	const initial = await walk(files, { followDynamic: false });
	assert.deepEqual(initial.files, ['entry.ts', 'helper.ts']);
	assert.deepEqual(initial.lazy, ['panel.ts']);
	assert.deepEqual(initial.bare, ['import-statement light <- helper.ts']);
});

test('minified ESM and CommonJS output are read', async () => {
	const { bare, files } = await walk({
		'entry.mjs': `import{a as b}from"min-a";export{c}from"./chunk.mjs";export const d=()=>import("min-c");console.log(b);`,
		'chunk.mjs': `export const c=require("min-cjs");`,
	});
	assert.deepEqual(files, ['entry.mjs', 'chunk.mjs']);
	assert.deepEqual(bare, [
		'dynamic-import min-c <- entry.mjs',
		'import-statement min-a <- entry.mjs',
		'require-call min-cjs <- chunk.mjs',
	]);
});

test('an unresolvable relative import fails loudly', async () => {
	await assert.rejects(walk({ 'entry.ts': `import './missing';` }), /missing/);
});
