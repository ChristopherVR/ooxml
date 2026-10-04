import { fileURLToPath } from 'node:url';

// Development resolves the core package from the repository root source, so UI tests never need
// a build of the core. Published consumers get the real dependency.
const coreSource = fileURLToPath(new URL('../../src', import.meta.url)).replaceAll('\\', '/');

export default {
	resolve: {
		alias: [{ find: /^ooxml-core\/(.+)$/, replacement: `${coreSource}/$1/index.ts` }],
	},
	test: {
		globals: true,
		// Component styles are `.css?raw` text; return the real CSS so tests can read it.
		css: { include: [/\.css\?raw$/] },
		environment: 'jsdom',
		include: ['src/**/*.test.ts'],
		exclude: ['node_modules', 'dist'],
	},
};
