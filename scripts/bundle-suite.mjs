/**
 * Bundle the suite host and its lazy editor chunks. Both Pages and desktop consume this output.
 * The static site contains no Office engine; the host imports the existing core and UI packages.
 */
import { createRequire } from 'node:module';
import { rm } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');

/** @param {string} outfile where to write the module */
export async function bundleSuite(outfile) {
	// Remove only the known generated chunk directory inside this build's output directory.
	const output = resolve(dirname(outfile));
	const chunks = resolve(output, 'suite-assets');
	if (dirname(chunks) !== output) throw new Error('Invalid generated asset directory');
	await rm(chunks, { recursive: true, force: true });
	const application = createRequire(join(ROOT, 'apps/office-suite/package.json'));
	const { build } = await import(pathToFileURL(application.resolve('vite')).href);
	await build({
		configFile: false,
		root: join(ROOT, 'apps/office-suite/src'),
		base: './',
		publicDir: false,
		worker: {
			format: 'es',
			rolldownOptions: {
				output: {
					entryFileNames: 'suite-assets/[name]-[hash].js',
					chunkFileNames: 'suite-assets/[name]-[hash].js',
				},
			},
		},
		resolve: {
			alias: [
				{
					find: /^pptx-vanilla-viewer$/,
					replacement: join(ROOT, 'viewers/pptx/packages/vanilla/src/viewer/PptxViewer.ts'),
				},
				{
					find: /^pptx-viewer-core$/,
					replacement: join(ROOT, 'viewers/pptx/packages/core/src/index.ts'),
				},
				{ find: /^ooxml-ui\/suite$/, replacement: join(ROOT, 'src/ui/src/suite/index.ts') },
				{ find: /^ooxml-core\/opc$/, replacement: join(ROOT, 'src/core/opc/index.ts') },
			],
		},
		build: {
			outDir: dirname(outfile),
			emptyOutDir: false,
			target: 'es2022',
			chunkSizeWarningLimit: 4000,
			rolldownOptions: {
				input: join(ROOT, 'apps/office-suite/src/main.js'),
				output: {
					entryFileNames: 'suite.js',
					chunkFileNames: 'suite-assets/[name]-[hash].js',
					assetFileNames: 'suite-assets/[name]-[hash][extname]',
				},
			},
		},
	});
}
