// node:test, not vitest: repo tooling tests run with `bun run test:scripts`.
import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { after, test } from 'node:test';

import {
	assertNoNodeOnlyModules,
	browserEntries,
	findNodeOnlyModules,
	isNodeOnlySpecifier,
} from './browser-bundle-guard.mjs';

const roots = [];
after(() => {
	for (const root of roots) {
		rmSync(root, { recursive: true, force: true });
	}
});

/** A consumer directory whose node_modules holds the given packages ({ name: { file: text } }). */
function consumer(packages) {
	const root = mkdtempSync(join(tmpdir(), 'browser-bundle-guard-'));
	roots.push(root);
	for (const [name, files] of Object.entries(packages)) {
		for (const [file, text] of Object.entries(files)) {
			const path = join(root, 'node_modules', name, file);
			mkdirSync(dirname(path), { recursive: true });
			writeFileSync(path, typeof text === 'string' ? text : JSON.stringify(text));
		}
	}
	return root;
}

function manifest(name, extra = {}) {
	return { name, type: 'module', exports: { '.': './index.js' }, ...extra };
}

async function reasons(root, packages) {
	const { findings } = await findNodeOnlyModules({ root, packages });
	return findings.map(({ entry, reason }) => `${entry}: ${reason}`).sort();
}

test('recognises Node built-ins and Node-only packages', () => {
	for (const specifier of ['fs', 'node:fs', 'fs/promises', 'node:worker_threads', 'module']) {
		assert.equal(isNodeOnlySpecifier(specifier), true, specifier);
	}
	for (const specifier of ['@napi-rs/canvas', 'canvas', 'jsdom', 'sharp', 'sharp/lib/index.js']) {
		assert.equal(isNodeOnlySpecifier(specifier), true, specifier);
	}
	// node-forge runs in browsers too.
	for (const specifier of [
		'react',
		'canvas-confetti',
		'pathe',
		'./fs',
		'ooxml-core/pptx',
		'node-forge',
	]) {
		assert.equal(isNodeOnlySpecifier(specifier), false, specifier);
	}
});

test('passes a package that reaches no Node-only module', async () => {
	const root = consumer({
		clean: {
			'package.json': manifest('clean'),
			// A literal import, a method named import and a message that mentions import() are fine.
			'index.js': [
				"export const lazy = () => import('./lazy.js');",
				'export class Sources { async import(custom) { return custom; } }',
				"export const hint = 'lazy(() => import(name)) expects a module';",
			].join('\n'),
			'lazy.js': 'export const value = 1;\n',
		},
	});
	assert.deepEqual(await reasons(root, ['clean']), []);
});

test('reports static, dynamic and nested imports of Node-only modules', async () => {
	const root = consumer({
		leaky: {
			'package.json': manifest('leaky'),
			'index.js': "import { readFileSync } from 'node:fs';\nexport { readFileSync };\n",
			'lazy.js':
				"export async function load() {\n\ttry { return await import('@napi-rs/canvas'); } catch { return undefined; }\n}\n",
		},
		host: {
			'package.json': manifest('host', { exports: { '.': './index.js', './lazy': './lazy.js' } }),
			'index.js': "export * from 'leaky';\n",
			'lazy.js': "export { load } from 'leaky/lazy.js';\n",
		},
	});
	writeFileSync(
		join(root, 'node_modules/leaky/package.json'),
		JSON.stringify(manifest('leaky', { exports: { '.': './index.js', './lazy.js': './lazy.js' } })),
	);
	assert.deepEqual(await reasons(root, ['host']), [
		'host/lazy: dynamic-import of @napi-rs/canvas from node_modules/leaky/lazy.js',
		'host: import-statement of node:fs from node_modules/leaky/index.js',
	]);
});

test('reports a dynamic import whose specifier a bundler cannot see', async () => {
	const root = consumer({
		opaque: {
			'package.json': manifest('opaque'),
			'index.js':
				"const specifier = ['@napi-rs', 'canvas'].join('/');\nexport const load = () => import(/* webpackIgnore: true */ specifier);\n",
		},
	});
	assert.deepEqual(await reasons(root, ['opaque']), [
		'opaque: dynamic import with a computed specifier at node_modules/opaque/index.js:2',
	]);
});

test('leaves optional peers the host has not installed to the host', async () => {
	const root = consumer({
		peer: {
			'package.json': manifest('peer'),
			'index.js': "export const ai = () => import('optional-ai-sdk');\n",
		},
	});
	assert.deepEqual(await reasons(root, ['peer']), []);
});

test('lists browser entries, expanding wildcards and skipping Node-only subpaths and assets', () => {
	const root = consumer({
		lib: {
			'package.json': manifest('lib', {
				exports: {
					'.': { types: './index.d.ts', import: './index.js' },
					'./cli': './cli.js',
					'./signature-node': './signature-node.js',
					'./node': './node.js',
					'./automation/node': './automation-node.js',
					'./styles.css': './styles.css',
					'./package.json': './package.json',
					'./editor/*': { import: './dist/editor/*.mjs' },
				},
			}),
			'dist/editor/a.mjs': '',
			'dist/editor/nested/b.mjs': '',
			'dist/editor/a.d.mts': '',
			'dist/editor/chunk-ABCD1234.mjs': '',
			'dist/editor/nested/chunk-x.mjs': '',
			'dist/editor/shared-ABCD1234.mjs': '',
			'dist/editor/chunked-view.mjs': '',
		},
	});
	assert.deepEqual(browserEntries(root, 'lib').sort(), [
		'lib',
		'lib/editor/a',
		'lib/editor/chunked-view',
		'lib/editor/nested/b',
	]);
});

test('assertNoNodeOnlyModules lists every finding in its error', async () => {
	const root = consumer({
		leaky: {
			'package.json': manifest('leaky'),
			'index.js': "import { readFileSync } from 'node:fs';\nexport { readFileSync };\n",
		},
		clean: { 'package.json': manifest('clean'), 'index.js': 'export const value = 1;\n' },
	});
	assert.equal(await assertNoNodeOnlyModules({ root, packages: ['clean'] }), 1);
	await assert.rejects(
		assertNoNodeOnlyModules({ root, packages: ['clean', 'leaky'] }),
		/Browser entries reach Node-only modules:\n {2}leaky: import-statement of node:fs/,
	);
});
