import { defineConfig } from 'vitest/config';
import { fileURLToPath } from 'node:url';
import { svelte } from '@sveltejs/vite-plugin-svelte';
import { localCoreAliases } from './scripts/local-core-aliases';

const local = (path: string) => fileURLToPath(new URL(path, import.meta.url));

export default defineConfig({
	// Compiles `XlsxEditor.svelte` for the Svelte binding contract test (client build, no HMR).
	plugins: [svelte({ hot: false, compilerOptions: { css: 'injected' } })],
	resolve: {
		// Svelte 5's `mount` exists only in its browser build.
		conditions: ['browser'],
		alias: [
			...localCoreAliases(local('.')),
			{ find: /^solid-js\/web$/, replacement: local('./node_modules/solid-js/web/dist/web.js') },
			{ find: /^solid-js$/, replacement: local('./node_modules/solid-js/dist/solid.js') },
			{
				find: /^@christophervr\/xlsx-core\/load$/,
				replacement: local('./packages/core/src/load.ts'),
			},
			{ find: /^@christophervr\/xlsx-core$/, replacement: local('./packages/core/src/index.ts') },
			{
				find: /^xlsx-bindings\/(react|vue|angular|solid|common)$/,
				replacement: local('./packages/bindings/src/$1'),
			},
			...['web-component', 'bindings'].map((name) => ({
				find: new RegExp(`^xlsx-${name}$`),
				replacement: local(`./packages/${name}/src/index.ts`),
			})),
		],
	},
	test: {
		include: ['packages/**/*.test.ts', 'packages/**/*.test.tsx', 'scripts/**/*.test.ts'],
		environment: 'node',
	},
});
