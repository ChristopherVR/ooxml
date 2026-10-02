import { defineConfig } from 'tsup';

/** Fast local build for the format's consumer, independent of the PowerPoint bundle. */
export default defineConfig({
	entry: { 'visio/index': 'src/visio/index.ts' },
	outDir: 'dist',
	tsconfig: 'tsconfig.build.json',
	format: ['esm', 'cjs'],
	outExtension: ({ format }) => ({ js: format === 'esm' ? '.mjs' : '.cjs' }),
	dts: true,
	splitting: false,
	sourcemap: false,
	clean: false,
	external: ['jszip', '@xmldom/xmldom'],
	treeshake: true,
	platform: 'neutral',
});
