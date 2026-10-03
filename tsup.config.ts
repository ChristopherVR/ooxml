import { defineConfig } from 'tsup';

// Every strict area (the shared building blocks and the docx format) is bundled the same way as
// the pptx area (see tsup.pptx.config.ts): ESM (`.mjs`) and CJS (`.cjs`) per subpath entry.
// Declarations come from `tsc` (tsconfig.build.json, emitDeclarationOnly).
export default defineConfig({
	entry: {
		index: 'src/index.ts',
		'units/index': 'src/units/index.ts',
		'color/index': 'src/color/index.ts',
		'chart/index': 'src/chart/index.ts',
		'text/index': 'src/text/index.ts',
		'geometry/index': 'src/geometry/index.ts',
		'xml/index': 'src/xml/index.ts',
		'opc/index': 'src/opc/index.ts',
		'diagram/index': 'src/diagram/index.ts',
		'digest/index': 'src/digest/index.ts',
		'math/index': 'src/math/index.ts',
		'docx/index': 'src/docx/index.ts',
		'docx/embedded': 'src/docx/embedded.ts',
		'docx/layout/index': 'src/docx/layout/index.ts',
		'docx/load/index': 'src/docx/load/index.ts',
		'collab/index': 'src/collab/index.ts',
		'visio/index': 'src/visio/index.ts',
		'xlsx/index': 'src/xlsx/index.ts',
		'xlsx/load/index': 'src/xlsx/load/index.ts',
	},
	outDir: 'dist',
	tsconfig: 'tsconfig.build.json',
	format: ['esm', 'cjs'],
	outExtension: ({ format }) => ({ js: format === 'esm' ? '.mjs' : '.cjs' }),
	dts: false,
	splitting: false,
	sourcemap: false,
	// dist also holds the declarations and the pptx bundle: never clean it here.
	clean: false,
	external: ['jszip', '@xmldom/xmldom', 'yjs', 'y-protocols', 'lib0'],
	treeshake: true,
	platform: 'neutral',
	// The legacy .doc and .xls loaders inline the shared ole2 codecs (a devDependency), like the pptx bundle.
	noExternal: [/^@christophervr\/ole2(?:\/|$)/],
});
