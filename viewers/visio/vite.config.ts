import { defineConfig } from 'vite';
import { resolve } from 'node:path';
export default defineConfig({
	base: './',
	build: {
		outDir: 'site-dist',
		rollupOptions: {
			input: {
				home: resolve('index.html'),
				demo: resolve('demo/index.html'),
				docs: resolve('docs/index.html'),
				parity: resolve('docs/parity.html'),
				architecture: resolve('docs/architecture.html'),
			},
		},
	},
});
