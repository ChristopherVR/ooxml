import { sep } from 'node:path';
import { fileURLToPath } from 'node:url';

// Development resolves the core package from the repository root source, so UI tests never need
// a build of the core. Published consumers get the real dependency.
const coreSource = fileURLToPath(new URL('../core', import.meta.url)).replaceAll(sep, '/');
const uiSource = fileURLToPath(new URL('./src', import.meta.url)).replaceAll(sep, '/');

export default {
	resolve: {
		alias: [
			{ find: /^ooxml-core\/(.+)$/, replacement: `${coreSource}/$1/index.ts` },
			// The package imports its own subpaths by name; resolve them to source, not to dist.
			{ find: /^ooxml-ui\/pptx\/dom$/, replacement: `${uiSource}/pptx/dom/index.ts` },
		],
	},
	test: {
		globals: true,
		// Editor-mounting tests build a full ProseMirror view; under a parallel run they can take 5s+.
		testTimeout: 20_000,
		// Component styles are `.css?raw` text; return the real CSS so tests can read it.
		css: { include: [/\.css\?raw$/] },
		environment: 'jsdom',
		setupFiles: ['./vitest.setup.ts'],
		include: ['src/**/*.test.ts'],
		exclude: ['node_modules', 'dist', 'src/pptx/**'],
	},
};
