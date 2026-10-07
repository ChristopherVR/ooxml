import { defineConfig, type Options } from 'tsup';

// The pptx area is bundled separately from the tsc build of the other areas: consumers
// (pptx-viewer-core and the CLI `bin`) need dual ESM/CJS output with the legacy Office codecs
// (`@christophervr/ole2`) inlined, which a plain tsc build cannot produce.
const shared: Options = {
	entry: {
		'pptx/automation/schemas/chart-formatting-schemas':
			'pptx/automation/schemas/chart-formatting-schemas.ts',
		'pptx/automation/schemas/chart-schemas': 'pptx/automation/schemas/chart-schemas.ts',
		'pptx/automation/schemas/chart-user-shape-schemas':
			'pptx/automation/schemas/chart-user-shape-schemas.ts',
		'pptx/automation/schemas/element-schemas': 'pptx/automation/schemas/element-schemas.ts',
		'pptx/automation/schemas/export-schemas': 'pptx/automation/schemas/export-schemas.ts',
		'pptx/automation/schemas/geometry-schemas': 'pptx/automation/schemas/geometry-schemas.ts',
		'pptx/automation/schemas/hyperlink-schemas': 'pptx/automation/schemas/hyperlink-schemas.ts',
		'pptx/automation/schemas/index': 'pptx/automation/schemas/index.ts',
		'pptx/automation/schemas/json-schemas': 'pptx/automation/schemas/json-schemas.ts',
		'pptx/automation/schemas/layout-schemas': 'pptx/automation/schemas/layout-schemas.ts',
		'pptx/automation/schemas/lock-schemas': 'pptx/automation/schemas/lock-schemas.ts',
		'pptx/automation/schemas/metadata-schemas': 'pptx/automation/schemas/metadata-schemas.ts',
		'pptx/automation/schemas/ole-schemas': 'pptx/automation/schemas/ole-schemas.ts',
		'pptx/automation/schemas/presentation-schemas':
			'pptx/automation/schemas/presentation-schemas.ts',
		'pptx/automation/schemas/section-schemas': 'pptx/automation/schemas/section-schemas.ts',
		'pptx/automation/schemas/slide-schemas': 'pptx/automation/schemas/slide-schemas.ts',
		'pptx/automation/schemas/smartart-schemas': 'pptx/automation/schemas/smartart-schemas.ts',
		'pptx/automation/schemas/table-style-schemas': 'pptx/automation/schemas/table-style-schemas.ts',
		'pptx/automation/schemas/template-schemas': 'pptx/automation/schemas/template-schemas.ts',
		'pptx/automation/schemas/theme-schemas': 'pptx/automation/schemas/theme-schemas.ts',
		'pptx/automation/schemas/validation-schemas': 'pptx/automation/schemas/validation-schemas.ts',

		'pptx/automation/index': 'pptx/automation/index.ts',
		'pptx/ui/index': 'pptx/ui/index.ts',
		'pptx/index': 'pptx/index.ts',
		'pptx/converter/index': 'pptx/converter/index.ts',
		'pptx/cli/index': 'pptx/cli/index.ts',
		'pptx/signature-node/index': 'pptx/signature-node/index.ts',
		'pptx/smartart-layouts/index': 'pptx/smartart-layouts/index.ts',
	},
	outDir: 'dist',
	tsconfig: 'tsconfig.pptx.json',
	outExtension: ({ format }) => ({ js: format === 'esm' ? '.mjs' : '.cjs' }),
	dts: false,
	splitting: false,
	sourcemap: false,
	// dist also holds the tsc output of the other areas: never clean it here.
	clean: false,
	external: [
		/^ooxml-core(?:\/|$)/,
		'yjs',
		'emf-converter',
		'mtx-decompressor',
		'jszip',
		'fast-xml-parser',
		'fflate',
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
};

// The ESM build splits shared code into chunks: without it every subpath entry (for example
// `pptx` and `pptx/automation`) carries its own copy of the engine, and a consumer that imports
// both ships it twice. The CJS build keeps one self-contained file per entry.
export default defineConfig([
	{ ...shared, format: ['esm'], splitting: true },
	{ ...shared, format: ['cjs'] },
]);
