import { defineConfig } from 'vite';
import { resolve } from 'node:path';

// The site is VitePress (docs/); Vite builds only the playground at /demo/ and the browser-test API.
// scripts/build-pages.mjs sets PAGES_OUT so the playground lands inside the documentation output.
export default defineConfig(({ mode }) => ({
	base: './',
	// The playground lives in demos/visio (the demos of every viewer are at the repository root); the
	// page keeps its demo/index.html output path, and the config's own folder keeps `public` and the output.
	root: resolve('../../demos/visio'),
	publicDir: resolve('public'),
	build: {
		outDir: resolve(
			process.env.PAGES_OUT ?? (mode === 'browser-tests' ? '.browser-test-dist' : 'site-dist'),
		),
		emptyOutDir: !process.env.PAGES_OUT,
		rollupOptions: {
			preserveEntrySignatures: 'strict',
			input: {
				...(mode === 'browser-tests'
					? { 'test-api': resolve('../../e2e/visio/browser-api.ts') }
					: {}),
				demo: resolve('../../demos/visio/demo/index.html'),
			},
			output: {
				entryFileNames: (chunk) =>
					chunk.name === 'test-api' ? 'test-api.js' : 'assets/[name]-[hash].js',
			},
		},
	},
}));
