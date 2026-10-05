import { defineConfig } from 'vite';
import { resolve } from 'node:path';

// The site is VitePress (docs/); Vite builds only the playground at /demo/ and the browser-test API.
// scripts/build-pages.mjs sets PAGES_OUT so the playground lands inside the documentation output.
export default defineConfig(({ mode }) => ({
	base: './',
	build: {
		outDir:
			process.env.PAGES_OUT ?? (mode === 'browser-tests' ? '.browser-test-dist' : 'site-dist'),
		emptyOutDir: !process.env.PAGES_OUT,
		rollupOptions: {
			preserveEntrySignatures: 'strict',
			input: {
				...(mode === 'browser-tests' ? { 'test-api': resolve('tests/browser-api.ts') } : {}),
				demo: resolve('demo/index.html'),
			},
			output: {
				entryFileNames: (chunk) =>
					chunk.name === 'test-api' ? 'test-api.js' : 'assets/[name]-[hash].js',
			},
		},
	},
}));
