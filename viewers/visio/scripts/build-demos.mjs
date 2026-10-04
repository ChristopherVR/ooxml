/**
 * Build the framework demos into site-dist/demo-<framework>/ for GitHub Pages. The vanilla demo is
 * the main site build's /demo/. Each framework page is generated from demo/index.html (one source
 * for the workspace shell) with its own entry, and built separately with only its framework's
 * plugins, as scripts/build-packages.mjs does for the packages.
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'vite';
import { FRAMEWORK_DEMOS } from './framework-demos.mjs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const bindings = resolve(root, 'packages/bindings');
const requireBinding = createRequire(resolve(bindings, 'package.json'));
const plugin = async (name) => import(pathToFileURL(requireBinding.resolve(name)).href);

/** `node scripts/build-demos.mjs [outDir]`: site-dist by default, the browser-test build otherwise. */
const outRoot = process.argv[2] ?? 'site-dist';
const shell = readFileSync(resolve(root, 'demo/index.html'), 'utf8');

for (const demo of FRAMEWORK_DEMOS) {
	const dir = resolve(bindings, 'demos', `demo-${demo.id}`);
	const page = shell
		.replace(
			'<title>Visio Viewer · Workspace</title>',
			`<title>Visio Viewer · ${demo.label} workspace</title>`,
		)
		.replace('href="./workspace.css"', 'href="../../../../demo/workspace.css"')
		.replace('src="./main.ts"', `src="./${demo.entry}"`);
	if (page === shell)
		throw new Error('demo/index.html no longer has the expected title, stylesheet or entry.');
	// Generated (git-ignored): the workspace shell has one source, demo/index.html.
	writeFileSync(resolve(dir, 'index.html'), page);
	const plugins =
		demo.id === 'solid'
			? [(await plugin('vite-plugin-solid')).default({ hot: false })]
			: demo.id === 'svelte'
				? [(await plugin('@sveltejs/vite-plugin-svelte')).svelte()]
				: [];
	const oxc =
		demo.id === 'angular'
			? { oxc: { decorator: { legacy: true } } }
			: demo.id === 'react'
				? { oxc: { jsx: { runtime: 'automatic', importSource: 'react' } } }
				: {};
	await build({
		configFile: false,
		logLevel: 'warn',
		root: dir,
		base: './',
		plugins,
		...oxc,
		resolve: { conditions: ['browser'] },
		server: { fs: { allow: [root] } },
		worker: { format: 'es' },
		// Each demo bundles its framework and the viewer, like the vanilla demo's chunks.
		build: {
			outDir: resolve(root, outRoot, `demo-${demo.id}`),
			emptyOutDir: true,
			chunkSizeWarningLimit: 4096,
		},
	});
	console.log(`Built the ${demo.label} demo into site-dist/demo-${demo.id}/.`);
}
