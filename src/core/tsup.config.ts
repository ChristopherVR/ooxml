import { existsSync } from 'node:fs';
import { dirname, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig, type Options } from 'tsup';
import { xlsxCjsLocaleImports, xlsxCommandLocales } from './scripts/xlsx-command-locales';

const pdfEntries = {
	'docx/pdf-convert/index': 'docx/pdf-convert/index.ts',
	'docx/pdf-import/index': 'docx/pdf-import/index.ts',
	'pdf/index': 'pdf/index.ts',
};
// The PDF writer and Word's Save as PDF are built apart from the reader, so the reader's entries keep
// only what they use. They import the areas they build on by name (`sibling-entries` below), so a
// host that already loads `ooxml-core/docx` or `ooxml-core/docx/layout` shares that copy, and the
// XML parser the writer reads SVG with when there is no DOMParser stays an entry loaded on first use.
const pdfWriteEntries = {
	'docx/pdf-export/index': 'docx/pdf-export/index.ts',
	'pdf/write/index': 'pdf/write/index.ts',
};
const coreRoot = dirname(fileURLToPath(import.meta.url));
/** The published areas a relative import in the PDF writer resolves to, by their index file. */
const SIBLING_ENTRIES = ['docx', 'docx/layout', 'drawingml', 'text', 'opc', 'geometry', 'digest'];
const siblingEntries: NonNullable<Options['esbuildPlugins']>[number] = {
	name: 'sibling-entries',
	setup(build) {
		build.onResolve({ filter: /^\./ }, (args) => {
			const target = resolve(args.resolveDir, args.path);
			const area = relative(coreRoot, target).split(sep).join('/');
			if (SIBLING_ENTRIES.some((name) => area === `${name}/index`))
				return { path: `ooxml-core/${area.slice(0, -'/index'.length)}`, external: true };
			// Word's Save as PDF uses the writer entry; the writer's own files stay in it.
			const importer = relative(coreRoot, args.importer).split(sep).join('/');
			if (
				area.startsWith('pdf/write/') &&
				!importer.startsWith('pdf/write/') &&
				existsSync(`${target}.ts`)
			)
				return { path: 'ooxml-core/pdf/write', external: true };
			return undefined;
		});
	},
};

/**
 * Word's layout draws charts with the chart area's renderer, which every Word editor also loads as
 * `ooxml-core/chart` (the chart commands and dialogs). The layout entry imports it by that name, so
 * a host shares one copy instead of carrying a second one inside the layout.
 */
const layoutSharedAreas: NonNullable<Options['esbuildPlugins']>[number] = {
	name: 'layout-shared-areas',
	setup(build) {
		build.onResolve({ filter: /^\./ }, (args) => {
			const area = relative(coreRoot, resolve(args.resolveDir, args.path)).split(sep).join('/');
			return area === 'chart/index' || area === 'chart'
				? { path: 'ooxml-core/chart', external: true }
				: undefined;
		});
	},
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
		'drawingml/inkml-runtime': 'drawingml/inkml/runtime.ts',
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
		// The docx entry loads mail merge by name on first use (`docx/merge-fields-loader.ts`).
		'ooxml-core/docx/mail-merge',
		'#embedded-workbook-chart',
		'#xpath-runtime',
		'#inkml-runtime',
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
	'automation/add-ins/index': addIns,
	'drawingml/index': drawingml,
	'docx/mail-merge/index': mailMerge,
	'docx/ui/index': wordUi,
	'math/index': math,
	'xlsx/load/index': xlsxLoad,
	'xlsx/collab/index': xlsxCollab,
	'visio/index': _visio,
	'visio/ui/index': _visioUi,
	'docx/index': docx,
	'docx/embedded': docxEmbedded,
	'docx/load/index': docxLoad,
	'opc/index': opc,
	'docx/layout/index': docxLayout,
	...eagerEntries
} = shared.entry;

// Word's model, embedded and loader entries each split their own ESM build, so a part only some
// documents need (the signature verifier) is a chunk loaded on first use instead of code every
// consumer downloads.
const ownChunks = (name: string, file: string, chunks: string): Options => ({
	...shared,
	entry: { [name]: file },
	format: ['esm'],
	splitting: true,
	esbuildOptions(options) {
		options.chunkNames = `${chunks}/[name]-[hash]`;
	},
});

export default defineConfig([
	{
		...shared,
		entry: {
			...eagerEntries,
			'docx/ui/references': 'docx/references/index.ts',
			'docx/ui/table-tools': 'docx/ui/table-tools.ts',
		},
		format: ['esm'],
		splitting: false,
	},
	{
		...shared,
		entry: { 'docx/layout/index': docxLayout },
		format: ['esm'],
		splitting: false,
		esbuildPlugins: [layoutSharedAreas],
	},
	{ ...shared, format: ['cjs'], splitting: false },
	ownChunks('docx/index', docx, 'docx/chunks'),
	ownChunks('docx/embedded', docxEmbedded, 'docx/embedded-chunks'),
	ownChunks('docx/load/index', docxLoad, 'docx/load/chunks'),
	// The package signer (`signature.signPackage`) downloads on the first signing, as a file of its
	// own beside the OPC entry, so reading and verifying packages does not carry it.
	{
		...shared,
		entry: { 'opc/index': opc, 'opc/sign-package': 'opc/signature/sign/sign-package.ts' },
		format: ['esm'],
		splitting: false,
		esbuildPlugins: [
			{
				name: 'opc-lazy-signer',
				setup(build) {
					build.onResolve({ filter: /^\.\/sign-package$/ }, (args) =>
						args.kind === 'dynamic-import'
							? { path: './sign-package.mjs', external: true }
							: undefined,
					);
				},
			},
		],
	},
	{
		...shared,
		entry: { 'automation/add-ins/index': addIns },
		format: ['esm'],
		splitting: true,
		esbuildOptions(options) {
			options.chunkNames = 'automation/add-ins/chunks/[name]-[hash]';
		},
	},
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
	// The equation editor's tools (`loadEquationTools`) download on first use. They import the
	// converters through the math entry (`./index`), so they share its copy instead of a second one.
	{
		...shared,
		entry: { 'math/index': math },
		format: ['esm'],
		splitting: false,
		esbuildPlugins: [
			{
				name: 'math-lazy-equation-tools',
				setup(build) {
					build.onResolve({ filter: /^\.\/equation-tools$/ }, (args) =>
						args.kind === 'dynamic-import'
							? { path: './equation-tools.mjs', external: true }
							: undefined,
					);
				},
			},
		],
	},
	{
		...shared,
		entry: { 'math/equation-tools': 'math/equation-tools.ts' },
		format: ['esm'],
		splitting: false,
		esbuildPlugins: [
			{
				name: 'math-tools-use-math-entry',
				setup(build) {
					build.onResolve({ filter: /^\.\/index$/ }, (args) =>
						/[\\/]math[\\/][^\\/]+$/.test(args.importer)
							? { path: './index.mjs', external: true }
							: undefined,
					);
				},
			},
		],
	},
	{
		...shared,
		entry: { 'docx/ui/index': wordUi },
		format: ['esm'],
		splitting: false,
		// Keep optional tools independently downloadable. Shared code can stay together in
		// the initial bundle instead of being split across several separately gzipped files.
		esbuildPlugins: [
			{
				name: 'word-ui-lazy-tools',
				setup(build) {
					build.onResolve({ filter: /^(?:\.\.\/references|\.\/table-tools)$/ }, (args) => {
						if (args.kind !== 'dynamic-import') return;
						return {
							path: args.path === '../references' ? './references.mjs' : './table-tools.mjs',
							external: true,
						};
					});
				},
			},
		],
		esbuildOptions(options) {
			options.minifySyntax = true;
		},
	},
	{
		...shared,
		entry: { 'drawingml/index': drawingml },
		format: ['esm'],
		splitting: true,
		esbuildOptions(options) {
			options.chunkNames = 'drawingml/chunks/[name]-[hash]';
			options.minifySyntax = true;
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
	{
		...shared,
		// The simplification codes, the standard fonts' encoding and widths and Word's page marks get
		// files of their own, so a product that labels its report with them, picks a standard font
		// before it exports (PowerPoint's PDF notice and text fonts) or draws page borders and line
		// numbers (Word's Print Layout) does not load the writer for them.
		entry: {
			...pdfWriteEntries,
			'pdf/write/notes': 'pdf/write/pdf-svg-notes.ts',
			'pdf/write/fonts': 'pdf/write/pdf-fonts.ts',
			'docx/pdf-export/page-marks': 'docx/pdf-export/page-marks.ts',
			// The export itself too, so the entry is only re-exports and a bundler can take the page
			// marks from it without the export.
			'docx/pdf-export/export': 'docx/pdf-export/export-pdf.ts',
		},
		format: ['esm'],
		splitting: true,
		external: [...shared.external, 'ooxml-core/xml'],
		esbuildPlugins: [siblingEntries],
		esbuildOptions(options) {
			options.chunkNames = 'pdf/write/chunks/[name]-[hash]';
		},
	},
	{
		...shared,
		entry: pdfWriteEntries,
		format: ['cjs'],
		splitting: false,
		external: [...shared.external, 'ooxml-core/xml'],
		esbuildPlugins: [siblingEntries],
	},
]);
