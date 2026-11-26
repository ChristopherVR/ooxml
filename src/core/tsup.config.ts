import { defineConfig, type Options } from 'tsup';
import { xlsxCjsLocaleImports, xlsxCommandLocales } from './scripts/xlsx-command-locales';

const pdfEntries = {
	'docx/pdf-convert/index': 'docx/pdf-convert/index.ts',
	'docx/pdf-import/index': 'docx/pdf-import/index.ts',
	'pdf/index': 'pdf/index.ts',
};

const xlsxEntries = {
	'xlsx/index': 'xlsx/index.ts',
	'xlsx/ui/index': 'xlsx/ui/index.ts',
	'xlsx/embedded-chart': 'xlsx/ui/dialogs/embedded-chart-entry.ts',
};
const visioEntries = {
	'visio/index': 'visio/index.ts',
	'visio/ui/index': 'visio/ui/index.ts',
};
const xlsxLocaleEntries = Object.fromEntries(
	['fr', 'de', 'es', 'zh-CN'].map((locale) => {
		return [`xlsx/ui/locales/xlsx-locale-${locale}`, `xlsx/ui/locales/${locale}/index.ts`];
	}),
);

// Every strict area (the shared building blocks and the docx format) is bundled the same way as
// the pptx area (see tsup.pptx.config.ts): ESM (`.mjs`) and CJS (`.cjs`) per subpath entry.
// Declarations come from `tsc` (tsconfig.build.json, emitDeclarationOnly).
const shared = {
	entry: {
		index: 'index.ts',
		'units/index': 'units/index.ts',
		'color/index': 'color/index.ts',
		'chart/index': 'chart/index.ts',
		'chart/region-map/index': 'chart/region-map/index.ts',
		'text/index': 'text/index.ts',
		'i18n/index': 'i18n/index.ts',
		'ai/index': 'ai/index.ts',
		'ribbon/index': 'ribbon/index.ts',
		'embedding/index': 'embedding/index.ts',
		'embedding/onedrive/index': 'embedding/onedrive/index.ts',
		'geometry/index': 'geometry/index.ts',
		'xml/index': 'xml/public.ts',
		'xml/xpath-runtime': 'xml/xpath-runtime.ts',
		'opc/index': 'opc/index.ts',
		'drawingml/index': 'drawingml/index.ts',
		'diagram/index': 'diagram/index.ts',
		'diagram/builtin/index': 'diagram/builtin/index.ts',
		'digest/index': 'digest/index.ts',
		'crypto/index': 'crypto/index.ts',
		'math/index': 'math/index.ts',
		'docx/index': 'docx/index.ts',
		'docx/embedded': 'docx/embedded.ts',
		'docx/layout/index': 'docx/layout/index.ts',
		'docx/load/index': 'docx/load/index.ts',
		'docx/mail-merge/index': 'docx/mail-merge/index.ts',
		'docx/ui/index': 'docx/ui/index.ts',
		'collab/index': 'collab/index.ts',
		'teams/index': 'teams/index.ts',
		...visioEntries,
		'visio/collab/index': 'visio/collab/index.ts',
		'xlsx/load/index': 'xlsx/load/index.ts',
		'xlsx/collab/index': 'xlsx/collab/index.ts',
		'automation/index': 'automation/index.ts',
		'automation/add-ins/index': 'automation/add-ins/index.ts',
		'automation/node': 'automation/node.ts',
	},
	outDir: 'dist',
	tsconfig: 'tsconfig.build.json',
	outExtension: ({ format }) => ({ js: format === 'esm' ? '.mjs' : '.cjs' }),
	dts: false,
	sourcemap: false,
	// dist also holds the declarations and the pptx bundle: never clean it here.
	clean: false,
	external: [
		'jszip',
		'@xmldom/xmldom',
		'yjs',
		'y-protocols',
		'lib0',
		'ooxml-core/math',
		'#embedded-workbook-chart',
		'#xpath-runtime',
	],
	treeshake: true,
	platform: 'neutral',
	// The legacy .doc and .xls loaders and the package encryption (its CFB container) inline the
	// shared ole2 codecs (a devDependency), like the pptx bundle.
	noExternal: [/^@christophervr\/ole2(?:\/|$)/],
} satisfies Options;

// The PDF reader and its converters split their ESM build, so the decoders a PDF rarely needs
// (JBIG2, JPEG 2000, the security handler) are chunks loaded on first use and the three entries
// share one copy of the reader. The CJS build keeps one self-contained file per entry.
const {
	index,
	'drawingml/index': drawingml,
	'docx/mail-merge/index': mailMerge,
	'docx/ui/index': wordUi,
	'xlsx/load/index': xlsxLoad,
	'xlsx/collab/index': xlsxCollab,
	'visio/index': _visio,
	'visio/ui/index': _visioUi,
	...eagerEntries
} = shared.entry;

export default defineConfig([
	{ ...shared, entry: eagerEntries, format: ['esm'], splitting: false },
	{ ...shared, format: ['cjs'], splitting: false },
	{
		...shared,
		// The root re-exports DrawingML. Keep its gallery assets lazy without pulling root
		// namespace dependencies into consumers of the DrawingML subpath.
		entry: { index },
		format: ['esm'],
		splitting: true,
		esbuildOptions(options) {
			options.chunkNames = 'chunks/root/[name]-[hash]';
		},
	},
	{
		...shared,
		entry: { 'docx/mail-merge/index': mailMerge },
		format: ['esm'],
		splitting: true,
		esbuildOptions(options) {
			options.chunkNames = 'docx/mail-merge/chunks/[name]-[hash]';
		},
	},
	{
		...shared,
		entry: { 'docx/ui/index': wordUi },
		format: ['esm'],
		splitting: true,
		esbuildOptions(options) {
			options.chunkNames = 'docx/ui/chunks/[name]-[hash]';
		},
	},
	{
		...shared,
		entry: { 'drawingml/index': drawingml },
		format: ['esm'],
		splitting: true,
		esbuildOptions(options) {
			options.chunkNames = 'drawingml/chunks/[name]-[hash]';
		},
	},
	// The viewer and parser share the Visio model in the browser.
	{
		...shared,
		entry: visioEntries,
		format: ['esm'],
		splitting: true,
		esbuildOptions(options) {
			options.chunkNames = 'visio/chunks/[name]-[hash]';
		},
	},
	// Keep the rich-data codecs on demand in both the loader and collaboration entries.
	{
		...shared,
		entry: { 'xlsx/load/index': xlsxLoad, 'xlsx/collab/index': xlsxCollab },
		format: ['esm'],
		splitting: true,
		esbuildOptions(options) {
			options.chunkNames = 'xlsx/loader-chunks/[name]-[hash]';
		},
	},
	// Share the workbook code between Excel's model and UI entries and keep
	// each non-English dictionary in its own dynamically imported module.
	{
		...shared,
		entry: { ...xlsxEntries, ...xlsxLocaleEntries },
		format: ['esm'],
		splitting: true,
		esbuildPlugins: [xlsxCommandLocales],
		esbuildOptions(options) {
			options.chunkNames = 'xlsx/chunks/[name]-[hash]';
		},
	},
	{
		...shared,
		entry: { ...xlsxEntries, ...xlsxLocaleEntries },
		format: ['cjs'],
		splitting: false,
		esbuildPlugins: [xlsxCommandLocales, xlsxCjsLocaleImports],
	},
	{
		...shared,
		entry: pdfEntries,
		format: ['esm'],
		splitting: true,
		esbuildOptions(options) {
			options.chunkNames = 'pdf/chunks/[name]-[hash]';
		},
	},
	{ ...shared, entry: pdfEntries, format: ['cjs'], splitting: false },
]);
