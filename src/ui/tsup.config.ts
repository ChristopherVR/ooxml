import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { defineConfig } from 'tsup';

// Component styles are real .css files imported as text (`import css from './x.css?raw'`), the
// same spelling Vite and Vitest use.
export const rawCss = {
	name: 'raw-css',
	setup(build: { onResolve: Function; onLoad: Function }) {
		build.onResolve({ filter: /\.css\?raw$/ }, (args: { resolveDir: string; path: string }) => ({
			// Keep the ?raw suffix: tsup's own CSS handling matches paths that end in .css.
			path: resolve(args.resolveDir, args.path),
			namespace: 'raw-css',
		}));
		build.onLoad({ filter: /.*/, namespace: 'raw-css' }, async (args: { path: string }) => ({
			contents:
				'export default ' +
				JSON.stringify(await readFile(args.path.replace(/\?raw$/, ''), 'utf8')) +
				';',
			loader: 'js',
		}));
	},
};

// ESM only: custom elements need a browser (or a DOM) anyway. Declarations come from tsc
// (tsconfig.build.json). The core package stays external: it is a real dependency.
export default defineConfig({
	entry: {
		index: 'src/index.ts',
		theme: 'src/theme.ts',
		icons: 'src/icons.ts',
		controls: 'src/controls.ts',
		presence: 'src/presence.ts',
		smartart: 'src/smartart.ts',
		teams: 'src/teams/index.ts',
		suite: 'src/suite/index.ts',
		xlsx: 'src/xlsx/index.ts',
		docx: 'src/docx/index.ts',
		visio: 'src/visio/index.ts',
		'edit-worker': 'src/visio/edit-worker.ts',
		'parse-worker': 'src/visio/parse-worker.ts',
		'clipboard-worker': 'src/visio/clipboard-worker.ts',
	},
	outDir: 'dist',
	tsconfig: 'tsconfig.build.json',
	format: ['esm'],
	dts: false,
	esbuildPlugins: [rawCss],
	// Shared helpers (registry, styles, icons) would otherwise be duplicated per entry, which
	// would break the idempotent registration across entries.
	splitting: true,
	clean: false,
	treeshake: true,
	platform: 'browser',
	external: [/^ooxml-core(\/|$)/, /^ooxml-ui\/pptx(\/|$)/],
});
