/**
 * Follows the runtime imports of a module and its relative dependencies, and reports the bare
 * specifiers the closure imports. esbuild parses each file, so this works on TypeScript and TSX
 * sources (imports used only as types are erased, as at runtime) and on built ESM or CJS output.
 * Static imports, re-exports, side-effect imports, `import()` with a literal specifier and
 * `require()` all count. Bare specifiers are reported, never resolved or followed.
 *
 * Used by `src/ui/src/pptx/ai/host.test.ts`, `src/ui/scripts/package-smoke.mjs`, the React
 * binding's optional-peer test and `viewers/pptx/scripts/check-binding-packages.mjs`.
 */
import { dirname, isAbsolute, resolve } from 'node:path';

import * as esbuild from 'esbuild';

/** Imports of files that are not modules: raw CSS, assets and other `?query` imports. */
const ASSET = /\?|\.(?:css|less|scss|svg|png|jpe?g|gif|webp|woff2?|ttf|otf|html|json)$/;

/**
 * Walk the runtime module graph from `entry`.
 *
 * @param {string} entry Path of the first module.
 * @param {{ followDynamic?: boolean }} [options] With `followDynamic: false`, relative
 *   `import()` targets (lazily loaded chunks) are listed in `lazy` instead of being walked, so
 *   `files` is the closure a bundler puts in the initial chunk. A CJS `require()` is always
 *   followed: a lazy `Promise.resolve().then(() => require('./chunk'))` is not told apart.
 * @returns {Promise<{ files: string[], lazy: string[], bare: Array<{ specifier: string, kind: string, importer: string }> }>}
 */
export async function walkModuleGraph(entry, options) {
	const [graph] = await walkModuleGraphs([entry], options);
	return graph;
}

/**
 * `walkModuleGraph` for several entries in one esbuild pass: a package's entries share most of
 * their graph, and parsing it once per entry was the bulk of the boundary tests' time.
 *
 * @param {string[]} entries Paths of the first modules.
 * @param {{ followDynamic?: boolean }} [options] As for `walkModuleGraph`.
 * @returns {Promise<Array<{ files: string[], lazy: string[], bare: Array<{ specifier: string, kind: string, importer: string }> }>>}
 *   One graph per entry, in order.
 */
export async function walkModuleGraphs(entries, { followDynamic = true } = {}) {
	const entryPaths = entries.map((entry) => resolve(entry));
	const workingDir = dirname(entryPaths[0]);
	const bare = [];
	const lazy = [];
	const plugin = {
		name: 'module-graph',
		setup(build) {
			build.onResolve({ filter: /.*/ }, async (args) => {
				if (args.kind === 'entry-point' || args.pluginData?.moduleGraph) {
					return undefined;
				}
				if (ASSET.test(args.path)) {
					return { path: args.path, external: true };
				}
				if (!args.path.startsWith('.') && !isAbsolute(args.path)) {
					bare.push({ specifier: args.path, kind: args.kind, importer: args.importer });
					return { path: args.path, external: true };
				}
				if (args.kind === 'dynamic-import' && !followDynamic) {
					const resolved = await build.resolve(args.path, {
						kind: args.kind,
						importer: args.importer,
						resolveDir: args.resolveDir,
						pluginData: { moduleGraph: true },
					});
					if (resolved.errors.length > 0) {
						return undefined;
					}
					lazy.push({ path: resolved.path, importer: args.importer });
					return { path: args.path, external: true };
				}
				return undefined;
			});
		},
	};
	const result = await esbuild.build({
		entryPoints: entryPaths,
		absWorkingDir: workingDir,
		outdir: resolve(workingDir, '.module-graph'),
		// Entries may share a file name (`public-api.ts` in two folders): keep their outputs apart.
		entryNames: '[dir]/[name]-[hash]',
		bundle: true,
		write: false,
		metafile: true,
		format: 'esm',
		platform: 'neutral',
		logLevel: 'silent',
		jsx: 'automatic',
		// Read every file as it is; a nearby tsconfig must not remap imports.
		tsconfigRaw: '{}',
		resolveExtensions: ['.ts', '.tsx', '.mts', '.mjs', '.js', '.cjs'],
		loader: { '.mjs': 'js', '.cjs': 'js' },
		plugins: [plugin],
	});
	const outputs = Object.values(result.metafile.outputs).filter((output) => output.entryPoint);
	return entryPaths.map((entryPath) => {
		const output = outputs.find((out) => resolve(workingDir, out.entryPoint) === entryPath);
		const files = Object.keys(output?.inputs ?? {}).map((input) => resolve(workingDir, input));
		const closure = new Set([entryPath, ...files]);
		return {
			files: [entryPath, ...files.filter((file) => file !== entryPath)],
			lazy: [
				...new Set(lazy.filter(({ importer }) => closure.has(importer)).map(({ path }) => path)),
			].filter((file) => !closure.has(file)),
			bare: bare.filter(({ importer }) => closure.has(importer)),
		};
	});
}
