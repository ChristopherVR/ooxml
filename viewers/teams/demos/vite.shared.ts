import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import type { UserConfig } from 'vite';

// Local development against a sibling checkout of the ooxml repository. By default a demo
// resolves ooxml-core and ooxml-ui from their SOURCE when that checkout exists (instant reload when
// you edit either; otherwise the published packages in node_modules), and the binding packages of
// this repository from their source too. TEAMS_USE_DIST=1 uses the built packages for both (run
// `bun run build:packages` first for this repository's own ones).
const ooxml = fileURLToPath(new URL(process.env.OOXML_DIR ?? '../../ooxml-core/', import.meta.url));
const useDist = process.env.TEAMS_USE_DIST === '1';
const ooxmlSource = !useDist && existsSync(`${ooxml}src/teams/index.ts`);
const local = (path: string): string =>
	fileURLToPath(new URL(`../packages/${path}`, import.meta.url));

export function sharedConfig(port: number): UserConfig {
	return {
		resolve: {
			// Yjs and Lit break (instanceof, double registration) if two copies load: force one.
			dedupe: [
				'yjs',
				'lib0',
				'y-protocols',
				'lit',
				'react',
				'react-dom',
				'vue',
				'solid-js',
				'svelte',
			],
			alias: [
				...(ooxmlSource
					? [
							{ find: /^ooxml-core\/(.+)$/, replacement: `${ooxml}src/$1/index.ts` },
							{ find: /^ooxml-ui$/, replacement: `${ooxml}packages/ui/src/index.ts` },
						]
					: []),
				// The bindings import the private web component by name; the build inlines it.
				...(useDist
					? []
					: [
							{ find: /^teams-viewer$/, replacement: local('web-component/src/index.ts') },
							{ find: /^openteams-react-viewer$/, replacement: local('react/src/index.ts') },
							{ find: /^openteams-vue-viewer$/, replacement: local('vue/src/index.ts') },
							{ find: /^openteams-solid-viewer$/, replacement: local('solid/src/index.ts') },
							{ find: /^openteams-angular-viewer$/, replacement: local('angular/src/index.ts') },
							{ find: /^openteams-svelte-viewer$/, replacement: local('svelte/src/Teams.svelte') },
						]),
			],
		},
		server: { port, strictPort: true, host: '127.0.0.1', fs: { allow: ['..', ooxml] } },
	};
}
