// Builds `dist/pptx-viewer-canvas.css`, published as `pptx-react-viewer/canvas.css`: the
// stylesheet of the read-only canvas alone. The Tailwind entry is `src/styles/pptx-viewer-canvas.css`;
// this script adds an `@source` line for every file `SlideCanvas` and `useViewerBuildingBlocks` can
// reach, read from an esbuild metafile of the real import graph (lazily loaded renderers included),
// so a new renderer is picked up without anyone editing a list. Runs after `build:css`, whose
// `dist/pptx-viewer.css` it checks against: every utility here must also be in the full sheet.
import { spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

import { build } from 'esbuild';

const PACKAGE_DIR = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const UI_SOURCE = resolve(PACKAGE_DIR, '../../../../src/ui/src');
const STYLES_DIR = join(PACKAGE_DIR, 'src/styles');
const ENTRY_CSS = join(STYLES_DIR, 'pptx-viewer-canvas.css');
const OUTPUT = join(PACKAGE_DIR, 'dist/pptx-viewer-canvas.css');
const FULL_OUTPUT = join(PACKAGE_DIR, 'dist/pptx-viewer.css');
const MARKER = '/* @canvas-sources */';

/** The read-only building blocks whose markup the canvas stylesheet covers. */
export const CANVAS_ENTRIES = [
	'src/viewer/components/SlideCanvas.tsx',
	'src/viewer/hooks/useViewerBuildingBlocks.ts',
];

const toPosix = (path) => path.split(sep).join('/');

/**
 * Picks the files whose class names Tailwind must scan from an esbuild metafile: this package's
 * sources, and the shared `ooxml-ui` render modules (the only `ooxml-ui` files the full stylesheet
 * scans). Only modules that contribute code after tree shaking count. Returns sorted absolute paths.
 */
export function selectSources(metafile, { packageDir = PACKAGE_DIR, uiSource = UI_SOURCE } = {}) {
	const used = new Set();
	for (const output of Object.values(metafile.outputs)) {
		for (const [input, { bytesInOutput }] of Object.entries(output.inputs)) {
			if (bytesInOutput > 0) {
				used.add(toPosix(resolve(packageDir, input)));
			}
		}
	}
	const own = `${toPosix(packageDir)}/src/`;
	const render = `${toPosix(uiSource)}/pptx/render/`;
	return [...used]
		.filter((file) => file.startsWith(own) || file.startsWith(render))
		.filter((file) => !/\.(test|spec)\.[cm]?[jt]sx?$/u.test(file))
		.sort();
}

/** Inserts one `@source` line per file (relative to the entry stylesheet) at the marker. */
export function canvasEntry(template, files, stylesDir = STYLES_DIR) {
	if (!template.includes(MARKER)) {
		throw new Error(`${ENTRY_CSS} lost its ${MARKER} marker`);
	}
	const lines = files.map((file) => `@source '${toPosix(relative(stylesDir, file))}';`);
	return template.replace(MARKER, lines.join('\n'));
}

/** Class selectors in the utilities layer of a minified Tailwind build. */
export function utilityClasses(css) {
	const start = css.indexOf('@layer utilities{');
	if (start === -1) {
		return new Set();
	}
	let depth = 0;
	let end = start;
	for (; end < css.length; end++) {
		if (css[end] === '{') {
			depth++;
		} else if (css[end] === '}' && --depth === 0) {
			break;
		}
	}
	// Drop the declaration blocks (always innermost) so a decimal value (`.5rem`) is not read as a
	// class; the braces of a nested `@media` stay, and so do the selectors inside them.
	const selectors = css.slice(start, end).replace(/\{[^{}]*\}/gu, ' ');
	return new Set(selectors.match(/\.(?:\\.|[\w-])+/gu) ?? []);
}

/** Bundles the canvas entries in memory and returns their metafile. */
export async function canvasMetafile() {
	const result = await build({
		absWorkingDir: PACKAGE_DIR,
		stdin: {
			contents: CANVAS_ENTRIES.map((entry, at) => `export * as e${at} from './${entry}';`).join(
				'\n',
			),
			resolveDir: PACKAGE_DIR,
			loader: 'ts',
			sourcefile: 'canvas-entry.ts',
		},
		bundle: true,
		splitting: true,
		format: 'esm',
		platform: 'browser',
		jsx: 'automatic',
		// Resolve through package exports only, never a tsconfig `paths` alias of a nearby project.
		tsconfigRaw: {},
		outdir: join(PACKAGE_DIR, 'dist/.canvas-graph'),
		write: false,
		metafile: true,
		logLevel: 'error',
		loader: { '.css': 'text' },
		// The shared UI is read from source, as the full stylesheet's `@source` does. The core is
		// bundled from its build so that the UI's `export *` of it stays statically analysable
		// (an external one would keep every UI module alive); every other package stays external.
		alias: { 'ooxml-ui': UI_SOURCE },
		plugins: [
			{
				name: 'canvas-graph-externals',
				setup(context) {
					context.onResolve({ filter: /\?/ }, (args) => ({ path: args.path, external: true }));
					context.onResolve({ filter: /^[^./]/ }, (args) =>
						/^(ooxml-ui|ooxml-core)(\/|$)/u.test(args.path)
							? undefined
							: { path: args.path, external: true },
					);
				},
			},
		],
	});
	return result.metafile;
}

function compile(entry) {
	const require = createRequire(import.meta.url);
	const cli = join(dirname(require.resolve('@tailwindcss/cli/package.json')), 'dist/index.mjs');
	const result = spawnSync(process.execPath, [cli, '-i', '-', '-o', OUTPUT, '--minify'], {
		cwd: STYLES_DIR,
		input: entry,
		encoding: 'utf8',
	});
	if (result.status !== 0) {
		throw new Error(`tailwindcss failed:\n${result.stderr}`);
	}
}

export async function buildCanvasCss() {
	const files = selectSources(await canvasMetafile());
	compile(canvasEntry(readFileSync(ENTRY_CSS, 'utf8'), files));
	const canvas = utilityClasses(readFileSync(OUTPUT, 'utf8'));
	const full = utilityClasses(readFileSync(FULL_OUTPUT, 'utf8'));
	const extra = [...canvas].filter((name) => !full.has(name));
	if (extra.length > 0) {
		throw new Error(`canvas.css has utilities styles.css lacks: ${extra.slice(0, 20).join(' ')}`);
	}
	return { files: files.length, bytes: readFileSync(OUTPUT).length };
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
	const { files, bytes } = await buildCanvasCss();
	writeFileSync(1, `canvas.css: ${bytes} bytes from ${files} source files\n`);
}
