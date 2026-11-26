import { defineConfig } from 'tsup';

/** Fast local build for the format's consumer, independent of the PowerPoint bundle. */
const visio = { 'visio/index': 'visio/index.ts', 'visio/ui/index': 'visio/ui/index.ts' };
const standalone = {
	'visio/collab/index': 'visio/collab/index.ts',
	'xlsx/embedded-chart': 'xlsx/ui/dialogs/embedded-chart-entry.ts',
	'xml/xpath-runtime': 'xml/xpath-runtime.ts',
};
const common = {
	outDir: 'dist',
	tsconfig: 'tsconfig.build.json',
	dts: false,
	sourcemap: false,
	clean: false,
	external: [
		'jszip',
		'@xmldom/xmldom',
		'yjs',
		'y-protocols',
		'lib0',
		'#embedded-workbook-chart',
		'#xpath-runtime',
	],
	treeshake: true,
	platform: 'neutral' as const,
};

export default defineConfig([
	{
		...common,
		entry: visio,
		format: ['esm'],
		splitting: true,
		outExtension: () => ({ js: '.mjs' }),
		esbuildOptions(options) {
			options.chunkNames = 'visio/chunks/[name]-[hash]';
		},
	},
	{
		...common,
		entry: standalone,
		format: ['esm'],
		splitting: false,
		outExtension: () => ({ js: '.mjs' }),
	},
	{
		...common,
		entry: { ...visio, ...standalone },
		format: ['cjs'],
		splitting: false,
		outExtension: () => ({ js: '.cjs' }),
	},
]);
