import { defineConfig } from 'tsup';

// Every strict area (the shared building blocks and the docx format) is bundled the same way as
// the pptx area (see tsup.pptx.config.ts): ESM (`.mjs`) and CJS (`.cjs`) per subpath entry.
// Declarations come from `tsc` (tsconfig.build.json, emitDeclarationOnly).
export default defineConfig({
	entry: {
		index: 'src/core/index.ts',
		'units/index': 'src/core/units/index.ts',
		'color/index': 'src/core/color/index.ts',
		'chart/index': 'src/core/chart/index.ts',
		'text/index': 'src/core/text/index.ts',
		'geometry/index': 'src/core/geometry/index.ts',
		'xml/index': 'src/core/xml/index.ts',
		'opc/index': 'src/core/opc/index.ts',
		'diagram/index': 'src/core/diagram/index.ts',
		'digest/index': 'src/core/digest/index.ts',
		'crypto/index': 'src/core/crypto/index.ts',
		'math/index': 'src/core/math/index.ts',
		'docx/index': 'src/core/docx/index.ts',
		'docx/embedded': 'src/core/docx/embedded.ts',
		'docx/layout/index': 'src/core/docx/layout/index.ts',
		'docx/load/index': 'src/core/docx/load/index.ts',
		'docx/ui/index': 'src/core/docx/ui/index.ts',
		'collab/index': 'src/core/collab/index.ts',
		'teams/index': 'src/core/teams/index.ts',
		'visio/index': 'src/core/visio/index.ts',
		'visio/ui/index': 'src/core/visio/ui/index.ts',
		'xlsx/index': 'src/core/xlsx/index.ts',
		'xlsx/load/index': 'src/core/xlsx/load/index.ts',
		'xlsx/ui/index': 'src/core/xlsx/ui/index.ts',
		'automation/index': 'src/core/automation/index.ts',
		'automation/node': 'src/core/automation/node.ts',
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
	// The legacy .doc and .xls loaders and the package encryption (its CFB container) inline the
	// shared ole2 codecs (a devDependency), like the pptx bundle.
	noExternal: [/^@christophervr\/ole2(?:\/|$)/],
});
