import { execFileSync } from 'node:child_process';
import { cpSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';
import { build } from 'vite';
import { VIEWER_PACKAGES } from './release-plan.mjs';

const root = resolve(import.meta.dirname, '..');
const bindingsRequire = createRequire(resolve(root, 'packages/bindings/package.json'));
const bindingPlugin = async (name) => import(pathToFileURL(bindingsRequire.resolve(name)).href);
process.chdir(root);
const temporary = resolve(root, '.package-build');
rmSync(temporary, { recursive: true, force: true });
mkdirSync(temporary);
writeFileSync(
	resolve(temporary, 'tsconfig.json'),
	JSON.stringify({
		extends: '../packages/bindings/tsconfig.json',
		compilerOptions: {
			noEmit: false,
			declaration: true,
			// Vite bundles every package's JavaScript; tsc only writes declarations.
			emitDeclarationOnly: true,
			rootDir: '..',
			outDir: './output',
			paths: { 'visio-core': ['../packages/core/src/index.ts'] },
		},
		include: [
			'../src/**/*.ts',
			'../packages/bindings/src/**/*.ts',
			'../packages/bindings/src/**/*.tsx',
			'../packages/*/src/index.ts',
		],
		exclude: ['../src/**/*.test.ts', '../packages/svelte/src/index.ts'],
	}),
);
execFileSync(
	process.execPath,
	[resolve(root, 'node_modules/typescript/bin/tsc'), '-p', resolve(temporary, 'tsconfig.json')],
	{ stdio: 'inherit' },
);

const declarations = resolve(temporary, 'output');
const copyDeclarations = (from, to) =>
	cpSync(from, to, {
		recursive: true,
		filter: (file) => !/\.(?:js|jsx)$/.test(file),
	});
for (const [key, meta] of Object.entries(VIEWER_PACKAGES)) {
	const directory = resolve(root, meta.dir);
	const outDir = resolve(directory, 'dist');
	rmSync(outDir, { recursive: true, force: true });
	mkdirSync(outDir);
	if (key === 'core') {
		const entry = "export * from 'ooxml-core/visio';\n";
		writeFileSync(resolve(outDir, 'index.js'), entry);
		writeFileSync(resolve(outDir, 'index.d.ts'), entry);
	} else {
		const plugins =
			key === 'solid'
				? [(await bindingPlugin('vite-plugin-solid')).default({ hot: false })]
				: key === 'svelte'
					? [(await bindingPlugin('@sveltejs/vite-plugin-svelte')).svelte()]
					: [];
		const entry = resolve(directory, 'src/index.ts');
		await build({
			configFile: false,
			root,
			base: './',
			plugins,
			// Angular's binding uses legacy decorators with inject(); Vite's Oxc transform lowers them.
			...(key === 'angular' ? { oxc: { decorator: { legacy: true } } } : {}),
			worker: { format: 'es' },
			build: {
				outDir,
				emptyOutDir: false,
				minify: false,
				lib: { entry, formats: ['es'], fileName: () => 'index.js' },
				rolldownOptions: {
					external: (id) =>
						/^(?:visio-core|ooxml-core|ooxml-ui|emf-converter|react|vue|@angular\/core|solid-js|svelte)(?:\/|$)/.test(
							id,
						),
				},
			},
		});
		if (key === 'solid' || key === 'svelte') {
			const serverPlugins =
				key === 'solid'
					? [(await bindingPlugin('vite-plugin-solid')).default({ hot: false, ssr: true })]
					: [(await bindingPlugin('@sveltejs/vite-plugin-svelte')).svelte()];
			await build({
				configFile: false,
				root,
				base: './',
				plugins: serverPlugins,
				worker: { format: 'es' },
				build: {
					ssr: entry,
					outDir,
					emptyOutDir: false,
					minify: false,
					rolldownOptions: {
						output: { entryFileNames: 'index.server.js' },
						external: (id) =>
							/^(?:visio-core|ooxml-core|ooxml-ui|emf-converter|solid-js|svelte)(?:\/|$)/.test(id),
					},
				},
			});
		}
		const types = resolve(outDir, 'types');
		copyDeclarations(resolve(declarations, 'src'), resolve(types, 'src'));
		const bindings = resolve(types, 'packages/bindings/src');
		mkdirSync(bindings, { recursive: true });
		cpSync(
			resolve(declarations, 'packages/bindings/src/common.d.ts'),
			resolve(bindings, 'common.d.ts'),
		);
		if (key === 'svelte') {
			cpSync(
				resolve(root, 'packages/bindings/src/VisioViewer.svelte'),
				resolve(outDir, 'VisioViewer.svelte'),
			);
			// Svelte compiles its component in the host project; its shared UI runtime is prebuilt.
			const component = readFileSync(resolve(outDir, 'VisioViewer.svelte'), 'utf8').replace(
				"'./common.js'",
				"'./runtime.js'",
			);
			writeFileSync(resolve(outDir, 'VisioViewer.svelte'), component);
			const runtimeEntry = resolve(temporary, 'svelte-runtime.ts');
			writeFileSync(
				runtimeEntry,
				"export * from '../src/index.js';\nexport * from '../packages/bindings/src/common.js';\nexport * from 'visio-core';\n",
			);
			await build({
				configFile: false,
				root,
				base: './',
				worker: { format: 'es' },
				build: {
					outDir,
					emptyOutDir: false,
					minify: false,
					lib: {
						entry: runtimeEntry,
						formats: ['es'],
						fileName: () => 'runtime.js',
					},
					rolldownOptions: {
						external: /^(?:visio-core|ooxml-core|ooxml-ui|emf-converter)(?:\/|$)/,
					},
				},
			});
			writeFileSync(
				resolve(outDir, 'runtime.d.ts'),
				"export * from './types/packages/bindings/src/common.js';\nexport * from './types/src/index.js';\nexport * from 'visio-core';\n",
			);
			writeFileSync(
				resolve(outDir, 'VisioViewer.svelte.d.ts'),
				`import type { Component } from 'svelte';
import type { ViewerProps, ViewerHandle } from './types/packages/bindings/src/common.js';
declare const VisioViewer: Component<ViewerProps & { class?: string; style?: string }, Pick<ViewerHandle, 'load' | 'replacePlainText' | 'applyEdits' | 'undo' | 'redo' | 'cancelEdit' | 'exportVsdx' | 'fit' | 'setLayerVisibility' | 'resetLayerVisibility' | 'exportSvg' | 'createPrintSnapshot'> & { getHandle(): ViewerHandle }>;
export default VisioViewer;
`,
			);
			writeFileSync(
				resolve(outDir, 'index.d.ts'),
				"export * from './types/src/index.js';\nexport * from 'visio-core';\nexport { default, default as VisioViewer } from './VisioViewer.svelte';\n",
			);
			writeFileSync(
				resolve(outDir, 'index.svelte.js'),
				"export * from './runtime.js';\nexport { default, default as VisioViewer } from './VisioViewer.svelte';\n",
			);
		} else {
			cpSync(
				resolve(declarations, `packages/bindings/src/${key}.d.ts`),
				resolve(bindings, `${key}.d.ts`),
			);
			copyDeclarations(resolve(declarations, meta.dir), resolve(types, meta.dir));
			writeFileSync(
				resolve(outDir, 'index.d.ts'),
				`export * from './types/${meta.dir}/src/index.js';\n`,
			);
		}
	}
	for (const name of ['LICENSE', 'NOTICE']) cpSync(resolve(root, name), resolve(directory, name));
}
console.log('Built seven installable Visio packages.');
