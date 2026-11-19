import { defineConfig } from 'vite';
import { fileURLToPath } from 'node:url';
import { svelte } from '@sveltejs/vite-plugin-svelte';
import tailwindcss from '@tailwindcss/vite';
import { angularAot } from '../../demos/angular-aot.mjs';
const demo = (path: string): string =>
	fileURLToPath(new URL(`../../demos/docx/${path}`, import.meta.url));

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
				composition: fileURLToPath(
					new URL('./../../demos/docx/demo-vanilla/building-blocks.html', import.meta.url),
				),
				editor: fileURLToPath(
					new URL('./../../demos/docx/demo-vanilla/index.html', import.meta.url),
				),
				collaboration: fileURLToPath(
					new URL('./../../demos/docx/demo-vanilla/collaboration.html', import.meta.url),
				),
			},
		},
	},
	resolve: {
		alias: [
			{
				find: /^docx-core\/embedded$/,
				replacement: fileURLToPath(new URL('./packages/core/src/embedded.ts', import.meta.url)),
			},
			...['core', 'web-component'].map((name) => ({
				find: new RegExp(`^docx-${name}$`),
				replacement: fileURLToPath(new URL(`./packages/${name}/src/index.ts`, import.meta.url)),
			})),
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
			'prosemirror-commands',
			'prosemirror-history',
			'prosemirror-keymap',
			'prosemirror-model',
			'prosemirror-state',
			'prosemirror-view',
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
