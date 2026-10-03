import { defineConfig } from 'vite';
import { resolve } from 'node:path';
export default defineConfig(({ mode }) => ({
	base: './',
	build: {
		outDir: mode === 'browser-tests' ? '.browser-test-dist' : 'site-dist',
		rollupOptions: {
			preserveEntrySignatures: 'strict',
			input: {
				...(mode === 'browser-tests' ? { 'test-api': resolve('tests/browser-api.ts') } : {}),
				home: resolve('index.html'),
				demo: resolve('demo/index.html'),
				docs: resolve('docs/index.html'),
				parity: resolve('docs/parity.html'),
				architecture: resolve('docs/architecture.html'),
			},
			output: {
				entryFileNames: (chunk) =>
					chunk.name === 'test-api' ? 'test-api.js' : 'assets/[name]-[hash].js',
			},
		},
	},
}));
