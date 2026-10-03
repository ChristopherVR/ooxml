import { defineConfig } from 'vite';
import { resolve } from 'node:path';

/**
 * Library build of the root package: Vite bundles the JavaScript (inlining `?inline` CSS and
 * emitting the parse and edit workers under `assets/`), as the framework packages and
 * docx-viewer do. `tsc` only emits the declarations.
 */
export default defineConfig({
	base: './',
	worker: { format: 'es' },
	build: {
		outDir: 'dist',
		emptyOutDir: true,
		minify: false,
		lib: { entry: resolve('src/index.ts'), formats: ['es'], fileName: () => 'index.js' },
		rolldownOptions: {
			external: (id) => /^(?:ooxml-core|ooxml-ui|emf-converter)(?:\/|$)/.test(id),
		},
	},
});
