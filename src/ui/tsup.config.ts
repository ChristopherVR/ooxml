import { readFile } from 'node:fs/promises';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'tsup';

const PACKAGE = dirname(fileURLToPath(import.meta.url));

// Component styles are real .css files imported as text (`import css from './x.css?raw'`), the
// same spelling Vite and Vitest use.
export const rawCss = {
	name: 'raw-css',
	setup(build: { onResolve: Function; onLoad: Function }) {
		build.onResolve({ filter: /\.css\?raw$/ }, (args: { resolveDir: string; path: string }) => {
			const file = resolve(args.resolveDir, args.path.replace(/\?raw$/, ''));
			return {
				// esbuild writes this path into the bundle as a comment, so it is relative: the
				// output then does not depend on (or reveal) where the package was built. Keep the
				// ?raw suffix: tsup's own CSS handling matches paths that end in .css.
				path: `${relative(PACKAGE, file).replaceAll('\\', '/')}?raw`,
				namespace: 'raw-css',
				pluginData: { file },
			};
		});
		build.onLoad(
			{ filter: /.*/, namespace: 'raw-css' },
			async (args: { pluginData: { file: string } }) => ({
				contents: `export default ${JSON.stringify(await readFile(args.pluginData.file, 'utf8'))};`,
				loader: 'js',
			}),
		);
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
		'chart-format-locales': 'src/chart/format-pane-locales/index.ts',
		presence: 'src/presence.ts',
		smartart: 'src/smartart.ts',
		teams: 'src/teams/index.ts',
		bindings: 'src/bindings/index.ts',
		'bindings/react': 'src/bindings/react.ts',
		'bindings/vue': 'src/bindings/vue.ts',
		'bindings/solid': 'src/bindings/solid.ts',
		suite: 'src/suite/index.ts',
		xlsx: 'src/xlsx/index.ts',
		docx: 'src/docx/index.ts',
		visio: 'src/visio/index.ts',
		'edit-worker': 'src/visio/edit-worker.ts',
		'parse-worker': 'src/visio/parse-worker.ts',
		'clipboard-worker': 'src/visio/clipboard-worker.ts',
		'stencil-worker': 'src/visio/stencil-worker.ts',
		'xlsx-open-worker': 'src/xlsx/xlsx-open-worker.ts',
		'xlsx-picture-compress-worker': 'src/xlsx/xlsx-picture-compress-worker.ts',
		'mailings-recipient-worker': 'src/docx/mailings/mailings-recipient-worker.ts',
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
