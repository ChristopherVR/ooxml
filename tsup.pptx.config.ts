import { defineConfig } from 'tsup';

// The pptx area is bundled separately from the tsc build of the other areas: consumers
// (pptx-viewer-core and the CLI `bin`) need dual ESM/CJS output with the legacy Office codecs
// (`@christophervr/ole2`) inlined, which a plain tsc build cannot produce.
export default defineConfig(() => ({
	entry: {
		'pptx/automation/index': 'src/core/pptx/automation/index.ts',
		'pptx/ui/index': 'src/core/pptx/ui/index.ts',
		'pptx/index': 'src/core/pptx/index.ts',
		'pptx/converter/index': 'src/core/pptx/converter/index.ts',
		'pptx/cli/index': 'src/core/pptx/cli/index.ts',
		'pptx/signature-node/index': 'src/core/pptx/signature-node/index.ts',
	},
	outDir: 'dist',
	tsconfig: 'tsconfig.pptx.json',
	format: ['esm', 'cjs'],
	outExtension: ({ format }) => ({ js: format === 'esm' ? '.mjs' : '.cjs' }),
	dts: false,
	splitting: false,
	sourcemap: false,
	// dist also holds the tsc output of the other areas: never clean it here.
	clean: false,
	external: [
		'yjs',
		'emf-converter',
		'mtx-decompressor',
		'jszip',
		'fast-xml-parser',
		'fs',
		'path',
		'node-forge',
		'xml-crypto',
		'@xmldom/xmldom',
		'@napi-rs/canvas',
		'crypto',
		'http',
		'https',
		'tls',
	],
	treeshake: true,
	platform: 'neutral',
	noExternal: [/^@christophervr\/ole2(?:\/|$)/],
}));
