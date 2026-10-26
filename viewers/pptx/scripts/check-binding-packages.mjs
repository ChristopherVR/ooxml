/**
 * Install the npm-packed UI bindings without workspace links and import each
 * public entry from an isolated Node consumer, then bundle every browser entry
 * of the bindings for the browser and fail if one reaches a Node-only module.
 * Run after all package builds.
 */
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { assertNoNodeOnlyModules } from '../../../scripts/browser-bundle-guard.mjs';
import { walkModuleGraph } from '../../../scripts/module-graph.mjs';
import { toPublishManifest, readWorkspacePackages } from './publish-manifest.mjs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const temporary = mkdtempSync(join(tmpdir(), 'pptx-binding-consumer-'));
const rootManifest = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
const angularHostPackages = [
	'@angular/common',
	'@angular/compiler',
	'@angular/core',
	'@angular/platform-browser',
].map((name) => `${name}@${rootManifest.dependencies[name]}`);
const packages = [
	{ key: 'core', name: 'pptx-viewer-core', entry: 'PptxHandler' },
	{ key: 'react', name: 'pptx-react-viewer', entry: 'PowerPointViewer' },
	{ key: 'vue', name: 'pptx-vue-viewer', entry: 'PowerPointViewer' },
	{ key: 'angular', name: 'pptx-angular-viewer', entry: 'PowerPointViewerComponent' },
	{ key: 'vanilla', name: 'pptx-vanilla-viewer', entry: 'createPptxViewer' },
	{ key: 'svelte', name: 'pptx-svelte-viewer', entry: 'PowerPointViewer' },
];
const bindings = [...packages.slice(1)];
// Bindings whose main entry loads the AI panel lazily, so their initial chunks must take the
// SDK-free helpers from ooxml-ui/pptx/ai/host, never the full ooxml-ui/pptx/ai barrel (zod, the
// automation catalogue, a dynamic import of the ai SDK). Angular defers the external runtime
// import from its flat FESM until the chat service starts.
const aiHostBindings = new Set(['react', 'vue', 'vanilla', 'svelte', 'angular']);

function run(command, args, cwd, capture = false) {
	const result = spawnSync(command, args, {
		cwd,
		encoding: 'utf8',
		stdio: capture ? 'pipe' : 'inherit',
		env: { ...process.env, NODE_PATH: '' },
	});
	if (result.error) {
		throw result.error;
	}
	if (result.status !== 0) {
		throw new Error(`${command} exited ${result.status}: ${result.stderr ?? ''}`);
	}
	return result.stdout;
}

function npm(args, cwd, capture = false) {
	if (process.platform === 'win32') {
		const npmCli = join(dirname(process.execPath), 'node_modules/npm/bin/npm-cli.js');
		return run(process.execPath, [npmCli, ...args], cwd, capture);
	}
	return run('npm', args, cwd, capture);
}

// `PptxI18nProvider` is exported from both React entries. Each entry is its own bundle, so a
// provider from one and a viewer from the other only work if they share the context module.
function sharedContextCheck(load) {
	return `const translate = (key) => (key === 'pptx.viewer.loading' ? 'Loading (host)' : undefined);
const pairs = [
  ['pptx-react-viewer', main.PptxI18nProvider, 'pptx-react-viewer/viewer', viewer.PowerPointViewer],
  ['pptx-react-viewer/viewer', viewer.PptxI18nProvider, 'pptx-react-viewer', main.PowerPointViewer],
];
for (const [providerFrom, Provider, viewerFrom, Viewer] of pairs) {
  const html = renderToString(createElement(Provider, { locale: 'de', translate }, createElement(Viewer, { content: null })));
  assert.ok(html.includes('Loading (host)'), 'a provider from ' + providerFrom + ' translates a viewer from ' + viewerFrom + ' (${load})');
}
console.log('pptx-react-viewer: provider and viewer from either entry share one i18n context (${load})');
`;
}

try {
	const tarballs = [];
	const cjsBindings = [];
	const rootEntries = [];
	// The bindings consume the core and UI built by this checkout. Registry
	// versions can lack exports introduced by an unreleased migration.
	const localDependencies = [
		{ key: 'ooxml-core', name: 'ooxml-core', sourceDir: resolve(root, '../../src/core') },
		{ key: 'ooxml-ui', name: 'ooxml-ui', sourceDir: resolve(root, '../../src/ui') },
	];
	for (const target of [...localDependencies, ...packages]) {
		const sourceDir = target.sourceDir ?? join(root, 'packages', target.key);
		const packDir = target.key === 'angular' ? join(sourceDir, 'dist') : sourceDir;
		const packed = JSON.parse(
			npm(
				[
					'pack',
					'--ignore-scripts',
					'--json',
					'--workspaces=false',
					'--pack-destination',
					temporary,
				],
				packDir,
				true,
			),
		)[0];
		const packedPaths = new Set(packed.files.map((file) => file.path));
		const shippedManifest = JSON.parse(readFileSync(join(packDir, 'package.json'), 'utf8'));
		const rootExport = shippedManifest.exports?.['.'];
		const entryFile =
			rootExport?.import ?? rootExport?.default ?? shippedManifest.module ?? shippedManifest.main;
		if (bindings.includes(target) && typeof rootExport?.require === 'string') {
			cjsBindings.push({ name: target.name, entry: target.entry });
		}
		if (aiHostBindings.has(target.key)) {
			// The ESM entry is what a web host bundles. A CJS build loads its lazy chunks with
			// `require()` inside a promise, which a module walk cannot tell from an eager require.
			rootEntries.push({ name: target.name, file: entryFile });
		}
		assert.ok(entryFile, `${target.name} has no declared root runtime entry`);
		assert.ok(
			packedPaths.has(entryFile.replace(/^\.\//u, '')),
			`${target.name} npm tarball is missing root entry ${entryFile}`,
		);

		// CI builds from workspace manifests, where publishable runtime workspace
		// dependencies still say workspace:*. Apply the exact release resolver to
		// the extracted npm tarball before repacking it as a consumer-installable
		// artifact. The packed file selection and JS/declaration files stay intact.
		const unpacked = join(temporary, `${target.key}-unpacked`);
		mkdirSync(unpacked);
		run('tar', ['-xzf', join(temporary, packed.filename), '-C', unpacked], temporary);
		const packageDir = join(unpacked, 'package');
		const manifestPath = join(packageDir, 'package.json');
		const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));
		if (target.key === 'angular') {
			assert.ok(
				manifest.peerDependencies?.['@angular/platform-browser'],
				'Angular package declares @angular/platform-browser as a host peer',
			);
		}
		writeFileSync(
			manifestPath,
			`${JSON.stringify(toPublishManifest(manifest, readWorkspacePackages()), null, '\t')}\n`,
		);
		const normalized = JSON.parse(
			npm(
				[
					'pack',
					'--ignore-scripts',
					'--json',
					'--workspaces=false',
					'--pack-destination',
					temporary,
				],
				packageDir,
				true,
			),
		)[0];
		tarballs.push(join(temporary, normalized.filename));
		console.log(`[binding-pack] ${target.name}: packed ${packed.files.length} files`);
	}

	const consumer = join(temporary, 'consumer');
	mkdirSync(consumer);
	writeFileSync(
		join(consumer, 'package.json'),
		JSON.stringify({
			name: 'pptx-binding-packed-consumer',
			private: true,
			type: 'module',
		}),
	);
	npm(
		[
			'install',
			'--ignore-scripts',
			'--omit=optional',
			'--no-audit',
			'--no-fund',
			'--package-lock=false',
			...angularHostPackages,
			...tarballs,
		],
		consumer,
	);
	const checks = JSON.stringify(bindings.map(({ name, entry }) => ({ name, entry })));
	writeFileSync(
		join(consumer, 'consumer.mjs'),
		`import assert from 'node:assert/strict';
const core = await import('pptx-viewer-core');
assert.equal(typeof core.PptxHandler, 'function', 'packed core exports PptxHandler');
await import('@angular/compiler');
const bindings = ${checks};
const failures = [];
for (const { name, entry } of bindings) {
  try {
    const api = await import(name);
    assert.ok(api[entry], name + ' exports ' + entry);
    console.log(name + ': isolated packed entry imported and exported ' + entry);
  } catch (error) {
    failures.push(name + ': ' + (error?.name ?? 'Error') + ': ' + (error?.message ?? String(error)));
  }
}
if (failures.length) {
  throw new AggregateError(failures.map((message) => new Error(message)), 'Packed binding imports failed');
}
console.log('All five packed binding entry points imported successfully.');
`,
	);
	run(process.execPath, ['consumer.mjs'], consumer);
	writeFileSync(
		join(consumer, 'consumer.cjs'),
		`const assert = require('node:assert/strict');
for (const { name, entry } of ${JSON.stringify(cjsBindings.map(({ name, entry }) => ({ name, entry })))}) {
  const api = require(name);
  assert.ok(api[entry], name + ' CommonJS entry exports ' + entry);
  console.log(name + ': isolated packed CommonJS entry imported and exported ' + entry);
}
`,
	);
	run(process.execPath, ['consumer.cjs'], consumer);

	writeFileSync(
		join(consumer, 'react-i18n-context.mjs'),
		`import assert from 'node:assert/strict';
import { createElement } from 'react';
import { renderToString } from 'react-dom/server';
import * as main from 'pptx-react-viewer';
import * as viewer from 'pptx-react-viewer/viewer';
${sharedContextCheck('ESM')}`,
	);
	run(process.execPath, ['react-i18n-context.mjs'], consumer);
	writeFileSync(
		join(consumer, 'react-i18n-context.cjs'),
		`const assert = require('node:assert/strict');
const { createElement } = require('react');
const { renderToString } = require('react-dom/server');
const main = require('pptx-react-viewer');
const viewer = require('pptx-react-viewer/viewer');
${sharedContextCheck('CommonJS')}`,
	);
	run(process.execPath, ['react-i18n-context.cjs'], consumer);

	for (const { name, file } of rootEntries) {
		const { files, bare } = await walkModuleGraph(join(consumer, 'node_modules', name, file), {
			followDynamic: false,
		});
		const eager = bare
			.filter(
				({ specifier, kind }) =>
					kind !== 'dynamic-import' &&
					/^(?:ooxml-ui\/pptx\/ai|ai|@ai-sdk\/.+|zod(?:\/.+)?|ooxml-core\/pptx\/automation(?:\/.+)?)$/.test(
						specifier,
					),
			)
			.map(({ specifier, kind, importer }) => `${kind} of ${specifier} in ${importer}`);
		assert.deepEqual(eager, [], `${name} ${file} loads the AI runtime in its initial chunks`);
		console.log(
			`[binding-pack] ${name} ${file}: ${files.length} initial modules, no ooxml-ui/pptx/ai or ai`,
		);
	}

	// A web host bundles the same packages for the browser, with no Node fallbacks. The core and
	// UI packages get the same check in their own package smoke tests.
	const browserEntries = await assertNoNodeOnlyModules({
		root: consumer,
		packages: packages.map(({ name }) => name),
	});
	console.log(`[binding-pack] ${browserEntries} browser entries reach no Node-only module`);
} finally {
	rmSync(temporary, { recursive: true, force: true });
}
