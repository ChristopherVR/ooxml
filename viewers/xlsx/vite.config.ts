import { defineConfig } from 'vite';
import { fileURLToPath } from 'node:url';
import { svelte } from '@sveltejs/vite-plugin-svelte';
import tailwindcss from '@tailwindcss/vite';
import { angularAot } from '../../demos/angular-aot.mjs';
import { localCoreAliases } from './scripts/local-core-aliases';
const demo = (path: string): string =>
	fileURLToPath(new URL(`../../demos/xlsx/${path}`, import.meta.url));

export default defineConfig({
	plugins: [
		tailwindcss(),
		svelte(),
		// The Angular modes are compiled ahead of time: the published site's script policy never
		// allows the run-time compiler. tsconfig.angular.json lists the Angular sources.
		await angularAot(demo('package.json'), demo('demo-vanilla/tsconfig.angular.json')),
	],
	// The Angular plugin turns Vite's TypeScript transform off unless the config sets one; every
	// other source still needs it.
	oxc: {},
	build: {
		target: 'es2022',
		rollupOptions: {
			input: {
				editor: fileURLToPath(
					new URL('./../../demos/xlsx/demo-vanilla/index.html', import.meta.url),
				),
			},
		},
	},
	resolve: {
		alias: [
			...localCoreAliases(fileURLToPath(new URL('.', import.meta.url))),
			{
				find: /^@christophervr\/xlsx-core\/load$/,
				replacement: fileURLToPath(new URL('./packages/core/src/load.ts', import.meta.url)),
			},
			{
				find: /^@christophervr\/xlsx-core$/,
				replacement: fileURLToPath(new URL('./packages/core/src/index.ts', import.meta.url)),
			},
			{
				find: /^xlsx-web-component$/,
				replacement: fileURLToPath(
					new URL('./packages/web-component/src/index.ts', import.meta.url),
				),
			},
		],
	},
	// Pre-bundle every framework the demo can mount on demand. Otherwise the first visit to a
	// framework makes Vite discover new dependencies, re-optimize and reload the page mid-session.
	optimizeDeps: {
		include: [
			'@angular/core',
			'@angular/common',
			'@angular/platform-browser',
			'@xmldom/xmldom',
			'jszip',
			'react',
			'react-dom/client',
			'rxjs',
			'solid-js',
			'solid-js/web',
			'svelte',
			'vue',
		],
	},
	server: { fs: { allow: [fileURLToPath(new URL('../../', import.meta.url))] } },
});
