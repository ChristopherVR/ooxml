/**
 * Install the npm-packed UI bindings without workspace links and import each
 * public entry from an isolated Node consumer, then bundle every browser entry
 * of the bindings for the browser and fail if one reaches a Node-only module.
 * Run after all package builds.
 */
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs';
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

	// `createAiAssistant` handles. Each binding keeps its handle registry in a module of its own
	// that no entry exports, so no published export can create or read a handle except
	// `createAiAssistant` (Angular's `/ai`, which ng-packagr builds on its own, reaches the root
	// through `createAiAssistantHandle`, which takes a config only). The viewer entries and `./ai`
	// must share that one module in every format the binding ships: a copy per entry would not
	// read the other's handles.
	const viewerEntries = {
		react: ['.', './viewer'],
		vue: ['.', './viewer'],
		svelte: ['.', './viewer'],
		vanilla: ['.'],
		angular: ['.', './internals'],
	};
	const conditionFile = (target, format) => {
		const picked =
			typeof target === 'string'
				? format === 'esm' && target
				: format === 'esm'
					? (target?.import ?? target?.default)
					: target?.require;
		const file = typeof picked === 'object' ? picked?.default : picked;
		return typeof file === 'string' && /\.[cm]?js$/u.test(file) ? file : undefined;
	};
	const REGISTRY_CALL = /\bcreateAiAssistantRegistry\b/u;
	const sources = new Map();
	const callsRegistry = (file) => {
		if (!sources.has(file)) {
			sources.set(file, REGISTRY_CALL.test(readFileSync(file, 'utf8')));
		}
		return sources.get(file);
	};
	const aiChecks = { esm: [], cjs: [] };
	for (const { key, name } of bindings) {
		const packageDir = join(consumer, 'node_modules', name);
		const { exports } = JSON.parse(readFileSync(join(packageDir, 'package.json'), 'utf8'));
		for (const format of ['esm', 'cjs']) {
			const entries = Object.entries(exports)
				.map(([subpath, target]) => [subpath, conditionFile(target, format)])
				.filter(([, file]) => file);
			const aiFile = Object.fromEntries(entries)['./ai'];
			if (!aiFile) {
				assert.equal(format, 'cjs', `${name} has an ESM ./ai entry`);
				continue;
			}
			const registryFiles = async (subpath, followDynamic) => {
				const file = join(packageDir, Object.fromEntries(entries)[subpath]);
				const graph = await walkModuleGraph(file, { followDynamic });
				return { graph, files: graph.files.filter(callsRegistry) };
			};
			// Every module of the format that creates a registry, from any entry, lazily or not.
			const everywhere = new Set();
			for (const [subpath] of entries) {
				for (const file of (await registryFiles(subpath, true)).files) {
					everywhere.add(file);
				}
			}
			assert.equal(everywhere.size, 1, `${name} (${format}) has one handle registry module`);
			const [registry] = everywhere;
			for (const subpath of viewerEntries[key]) {
				const { files } = await registryFiles(subpath, false);
				assert.deepEqual(files, [registry], `${name} ${subpath} loads the registry (${format})`);
			}
			const ai = await registryFiles('./ai', false);
			if (key === 'angular') {
				// ng-packagr entry points cannot share a file: `/ai` imports the root by name.
				assert.deepEqual(ai.files, [], `${name}/ai has no registry of its own`);
				assert.ok(
					ai.graph.bare.some(
						({ specifier, kind }) => specifier === name && kind === 'import-statement',
					),
					`${name}/ai imports ${name}`,
				);
				assert.equal(registry, join(packageDir, Object.fromEntries(entries)['.']));
			} else {
				assert.deepEqual(ai.files, [registry], `${name}/ai shares the registry (${format})`);
			}
			const relative = registry
				.slice(packageDir.length + 1)
				.split('\\')
				.join('/');
			console.log(`[binding-pack] ${name} (${format}): ./ai and the viewer share ${relative}`);
			aiChecks[format].push({
				key,
				name,
				entries: entries.map(([subpath]) => subpath),
				registry: key === 'angular' ? undefined : realpathSync(registry),
				root: realpathSync(join(packageDir, Object.fromEntries(entries)['.'])),
			});
		}
	}
	// At runtime: the handle the packed `/ai` creates resolves in the registry module the viewer
	// entries import, and no entry of a binding or of ooxml-ui exports a registry or any
	// handle function other than `createAiAssistant`, Angular's `createAiAssistantHandle` and
	// ooxml-ui's `createAiAssistantRegistry` (which makes a new registry no viewer reads).
	const aiEntriesCheck = (format, plan) => {
		const load = (specifier) =>
			format === 'esm' ? `(await import(${specifier}))` : `require(${specifier})`;
		return `const config = { connection: { kind: 'endpoint', api: '/api/chat' } };
const registriesIn = (api) => Object.values(api).filter(
  (value) => value && typeof value === 'object' && typeof value.register === 'function' && typeof value.resolve === 'function',
);
const handleExports = (api) => Object.keys(api).filter((key) => /AiAssistant/u.test(key)).sort();
for (const { key, name, entries, registry, root } of ${JSON.stringify(plan)}) {
  for (const subpath of entries) {
    const api = ${load('name + subpath.slice(1)')};
    const expected = subpath === './ai' ? ['createAiAssistant'] : key === 'angular' && api === ${load('name')} ? ['createAiAssistantHandle'] : [];
    assert.deepEqual(handleExports(api), expected, name + subpath.slice(1) + ' handle exports (${format})');
    assert.deepEqual(registriesIn(api), [], name + subpath.slice(1) + ' exports no handle registry (${format})');
  }
  const ai = ${load("name + '/ai'")};
  const handle = ai.createAiAssistant(config);
  assert.deepEqual(Reflect.ownKeys(handle), [], name + '/ai returns an opaque handle (${format})');
  assert.ok(Object.isFrozen(handle), name + '/ai returns a frozen handle (${format})');
  assert.equal(Object.prototype.toString.call(handle), '[object PptxAiAssistant]');
  if (registry) {
    const [shared, ...others] = registriesIn(${load(format === 'esm' ? 'pathToFileURL(registry).href' : 'registry')});
    assert.equal(others.length, 0, name + ' registry module holds one registry (${format})');
    const resolved = shared.resolve(handle, { development: false });
    assert.equal(resolved?.config, config, name + '/ai handle resolves in the viewer registry (${format})');
    assert.ok(resolved?.panel, name + '/ai handle carries its chat panel (${format})');
  } else {
    assert.equal(ai.createAiAssistant.length, 1);
    assert.equal(${load('name')}.createAiAssistantHandle.length, 1, name + ' createAiAssistantHandle takes a config only');
    ${format === 'esm' ? "assert.equal(realpathSync(fileURLToPath(import.meta.resolve(name))), root, name + '/ai imports the root entry');" : ''}
  }
  console.log(name + ': packed ./ai handle resolves through the viewer registry; no entry exports a registry (${format})');
}
for (const specifier of ['ooxml-ui/pptx', 'ooxml-ui/pptx/ai/host', 'ooxml-ui/pptx/ai']) {
  const api = ${load('specifier')};
  assert.deepEqual(handleExports(api), ['createAiAssistantRegistry'], specifier + ' handle exports (${format})');
  assert.deepEqual(registriesIn(api), [], specifier + ' exports no handle registry (${format})');
  const first = api.createAiAssistantRegistry({ owner: 'host' });
  assert.notEqual(api.createAiAssistantRegistry({ owner: 'host' }), first, specifier + ' creates a new registry per call');
}
console.log('ooxml-ui: entries export only the registry factory (${format})');
`;
	};
	writeFileSync(
		join(consumer, 'ai-entries.mjs'),
		`import assert from 'node:assert/strict';
import { realpathSync } from 'node:fs';
import { fileURLToPath, pathToFileURL } from 'node:url';
await import('@angular/compiler');
${aiEntriesCheck('esm', aiChecks.esm)}`,
	);
	run(process.execPath, ['ai-entries.mjs'], consumer);
	writeFileSync(
		join(consumer, 'ai-entries.cjs'),
		`const assert = require('node:assert/strict');
${aiEntriesCheck('cjs', aiChecks.cjs)}`,
	);
	run(process.execPath, ['ai-entries.cjs'], consumer);

	// Which packed entries reach the AI runtime: the binding's framework SDK and the full
	// ooxml-ui/pptx/ai barrel (zod, the automation catalogue, the `ai` SDK). Today every viewer
	// entry and `./ai` reach them through the same lazy panel loader, and never from their
	// initial chunks. The SDK-free main entries of a later major release expect only ['./ai']
	// in every list (docs/guide/ai-assistant.md, "What the next major release changes").
	const aiRuntimeReach = {
		react: {
			'@ai-sdk/react': ['.', './ai', './viewer'],
			'ooxml-ui/pptx/ai': ['.', './ai', './viewer'],
		},
		vue: {
			'@ai-sdk/vue': ['.', './ai', './viewer'],
			'ooxml-ui/pptx/ai': ['.', './ai', './viewer'],
		},
		svelte: {
			'@ai-sdk/svelte': ['.', './ai', './viewer'],
			'ooxml-ui/pptx/ai': ['.', './ai', './viewer'],
		},
		vanilla: { 'ooxml-ui/pptx/ai': ['.', './ai'] },
		// Angular's `./ai` holds no loader: it imports the package root by name, which the walk
		// does not follow, and the root's chat service loads the runtime for every `ai` value.
		angular: { 'ooxml-ui/pptx/ai': ['.', './internals'] },
	};
	for (const { key, name } of bindings) {
		const packageDir = join(consumer, 'node_modules', name);
		const { exports } = JSON.parse(readFileSync(join(packageDir, 'package.json'), 'utf8'));
		assert.ok(exports['./ai'], `${name} exports ./ai`);
		const reached = Object.fromEntries(Object.keys(aiRuntimeReach[key]).map((id) => [id, []]));
		for (const [subpath, target] of Object.entries(exports)) {
			const file = target?.import ?? target?.default;
			if (typeof file !== 'string' || !/\.m?js$/u.test(file)) {
				continue;
			}
			const entry = join(packageDir, file);
			const initial = (await walkModuleGraph(entry, { followDynamic: false })).bare;
			const all = (await walkModuleGraph(entry)).bare;
			for (const id of Object.keys(reached)) {
				assert.deepEqual(
					initial
						.filter(({ specifier, kind }) => specifier === id && kind !== 'dynamic-import')
						.map(({ importer }) => importer),
					[],
					`${name} ${subpath} imports ${id} statically`,
				);
				if (all.some(({ specifier }) => specifier === id)) {
					reached[id].push(subpath);
				}
			}
		}
		for (const [id, subpaths] of Object.entries(reached)) {
			assert.deepEqual(subpaths.sort(), aiRuntimeReach[key][id], `${name} entries reaching ${id}`);
			console.log(`[binding-pack] ${name} entries reaching ${id} (dynamically only): ${subpaths}`);
		}
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
