/**
 * Fails when a host bundler can no longer tree-shake the built React binding.
 *
 *   node scripts/check-tree-shaking.mjs [--json]     (or: bun run check:tree-shaking)
 *
 * Run after `bun run build:packages`. A host that renders only `SlideCanvas` from
 * `useViewerBuildingBlocks` used to download the whole editor, because the bundled `dist` kept
 * every component through module-level side effects (`forwardRef(...)`, `memo(...)`,
 * `createContext(...)`, `X.displayName = ...`) and through runtime namespace copies
 * (`__reExport(...)`) of `export *` re-exports. Each case below bundles one import of the built
 * `dist/index.mjs` with rollup, every bare import kept external and every module of the binding
 * treated as side-effectful (so only statement-level purity counts, as it does inside one of the
 * binding's chunks), and compares the bytes of the binding's own code that survive with a budget.
 * The budgets hold today's size with some headroom: raise one only for a deliberate addition.
 *
 * Tree shaking must not drop what the binding does on import, either. A host bundler (webpack,
 * Rspack, esbuild, rollup) may skip a whole module that the package's `sideEffects` field calls
 * side-effect-free, so a bare `import 'pptx-react-viewer'` is bundled once more with that field
 * honoured, as a host would, and every call in `IMPORT_TIME_EFFECTS` has to survive it.
 */
import { existsSync, readdirSync, readFileSync, realpathSync } from 'node:fs';
import { dirname, isAbsolute, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { rollup } from 'rollup';

const viewerRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

/** Each case: what a host imports, and the most bytes of the binding's code it may keep. */
export const TREE_SHAKING_CASES = [
	{
		name: 'side-effect import',
		code: `import ${JSON.stringify('<dist>/index.mjs')};`,
		// Only `defineThreeViewElement()` (a real registration) and a few constants remain.
		budget: 8_000,
	},
	{
		name: 'read-only canvas (SlideCanvas, useViewerBuildingBlocks)',
		code: `export { SlideCanvas, useViewerBuildingBlocks } from ${JSON.stringify('<dist>/index.mjs')};`,
		budget: 900_000,
	},
];

/** A bare side-effect import of the binding, as a host writes it. */
const BARE_IMPORT = `import ${JSON.stringify('<dist>/index.mjs')};`;

/**
 * What a bare import of the binding must still do, as the external binding each effect calls.
 * `defineThreeViewElement()` registers `<pptx-three-view>` (`ThreeView.tsx`); importing
 * `ooxml-ui/pptx` for it also runs that package's own import-time registrations.
 */
export const IMPORT_TIME_EFFECTS = [
	{
		source: 'ooxml-ui/pptx',
		name: 'defineThreeViewElement',
		effect: 'the <pptx-three-view> registration',
	},
];

/** Files of `dir` (recursively) whose ESM keeps a runtime copy of a star re-export. */
export function findNamespaceCopies(dir) {
	const found = [];
	for (const entry of readdirSync(dir, { withFileTypes: true })) {
		const path = join(dir, entry.name);
		if (entry.isDirectory()) {
			found.push(...findNamespaceCopies(path));
		} else if (entry.name.endsWith('.mjs') && readFileSync(path, 'utf8').includes('__reExport(')) {
			found.push(path);
		}
	}
	return found;
}

const GLOB_TOKENS = { '**/': '(?:.*/)?', '**': '.*', '*': '[^/]*', '?': '[^/]' };

/**
 * Whether a module of the package at `packageDir` may have side effects, by the package's
 * `sideEffects` field, read the way host bundlers read it: missing or `true` means every module,
 * `false` none, and a list of globs only the files it matches (a glob without `/` matches the
 * file name in any folder). Without a `package.json` every module counts as side-effectful.
 */
export function packageSideEffects(packageDir) {
	const manifest = join(packageDir, 'package.json');
	const { sideEffects } = existsSync(manifest) ? JSON.parse(readFileSync(manifest, 'utf8')) : {};
	if (sideEffects === undefined || sideEffects === true) {
		return () => true;
	}
	if (!Array.isArray(sideEffects)) {
		return () => false;
	}
	const patterns = sideEffects.map((glob) => {
		const body = glob
			.replace(/^\.\//, '')
			.replace(/[.+^${}()|[\]\\]/g, '\\$&')
			.replace(/\*\*\/|\*\*|\*|\?/g, (m) => GLOB_TOKENS[m]);
		return new RegExp(`^${glob.includes('/') ? '' : '(?:.*/)?'}${body}$`);
	});
	return (file) => {
		const path = relative(packageDir, file).replaceAll('\\', '/');
		return patterns.some((pattern) => pattern.test(path));
	};
}

/**
 * The chunks that load with `code` (with `<dist>` resolved), bundled with rollup, every bare
 * import external. `hasSideEffects(file)` decides for each module of the binding whether it is
 * kept even when nothing of it is used; lazy chunks behind `import()` calls are left out.
 */
async function loadedChunks(distDir, code, hasSideEffects) {
	const entryId = '\0tree-shaking-entry';
	const bundle = await rollup({
		input: entryId,
		logLevel: 'silent',
		treeshake: {
			moduleSideEffects: (id, external) => !external && (id === entryId || hasSideEffects(id)),
		},
		plugins: [
			{
				name: 'tree-shaking-entry',
				resolveId(source, importer) {
					if (source === entryId) {
						return entryId;
					}
					if (source.startsWith('.') && importer) {
						return resolve(dirname(importer), source);
					}
					if (isAbsolute(source)) {
						return source;
					}
					return { id: source, external: true };
				},
				load(id) {
					return id === entryId
						? code.replaceAll('<dist>', distDir.replaceAll('\\', '/'))
						: undefined;
				},
			},
		],
	});
	const { output } = await bundle.generate({ format: 'es', compact: true });
	await bundle.close();
	const chunks = new Map(output.filter((c) => c.type === 'chunk').map((c) => [c.fileName, c]));
	const loaded = new Set();
	const pending = [...chunks.values()].filter((c) => c.isEntry).map((c) => c.fileName);
	while (pending.length) {
		const file = pending.pop();
		if (loaded.has(file) || !chunks.has(file)) {
			continue;
		}
		loaded.add(file);
		pending.push(...chunks.get(file).imports);
	}
	return [...loaded].map((file) => chunks.get(file));
}

/** Bytes of the binding's own code that survive bundling `code` (with `<dist>` resolved). */
export async function retainedBytes(distDir, code) {
	let bytes = 0;
	for (const chunk of await loadedChunks(distDir, code, () => true)) {
		for (const module of Object.values(chunk.modules)) {
			bytes += module.renderedLength;
		}
	}
	return bytes;
}

/**
 * The `IMPORT_TIME_EFFECTS` (by `effect`) that a bare import of the binding in `distDir` loses
 * when its package's `sideEffects` field is honoured. The package is the parent of `distDir`.
 */
export async function droppedImportTimeEffects(distDir, effects = IMPORT_TIME_EFFECTS) {
	const chunks = await loadedChunks(distDir, BARE_IMPORT, packageSideEffects(dirname(distDir)));
	const kept = (source, name) =>
		chunks.some((chunk) => chunk.importedBindings[source]?.includes(name));
	return effects.filter(({ source, name }) => !kept(source, name)).map(({ effect }) => effect);
}

/** Every failed expectation for the built binding in `distDir` (empty when all hold). */
export async function checkTreeShaking(distDir = join(viewerRoot, 'packages/react/dist')) {
	const problems = [];
	const sizes = {};
	for (const file of findNamespaceCopies(distDir)) {
		problems.push(
			`${file} copies a namespace at runtime (__reExport): a star re-export of an external ` +
				'module was bundled, so a host keeps every export of that module. Re-export by name.',
		);
	}
	for (const { name, code, budget } of TREE_SHAKING_CASES) {
		const bytes = await retainedBytes(distDir, code);
		sizes[name] = { bytes, budget };
		if (bytes > budget) {
			problems.push(
				`${name}: ${bytes} bytes of pptx-react-viewer survive tree shaking (budget ${budget}). ` +
					'A module-level call or assignment keeps unused code: mark pure factories with a ' +
					'pure annotation, or move the side effect into the code that needs it.',
			);
		}
	}
	const dropped = await droppedImportTimeEffects(distDir);
	for (const effect of dropped) {
		problems.push(
			`side-effect import: ${effect} no longer runs on a bare import of pptx-react-viewer. ` +
				"Keep the call at module level, unannotated, in a module the package's " +
				'`sideEffects` field does not mark side-effect-free.',
		);
	}
	const kept = IMPORT_TIME_EFFECTS.map(({ effect }) => effect).filter((e) => !dropped.includes(e));
	return { problems, sizes, kept };
}

/**
 * Whether `argvPath` (normally `process.argv[1]`) names the module at `moduleUrl`. Both sides are
 * canonicalised, because Node reports `import.meta.url` with symlinks resolved while `argv[1]` is
 * the path as typed (on macOS `/tmp` is a link to `/private/tmp`): comparing them as given made
 * the check exit 0 without running when invoked through a symlinked path.
 */
export function isCliEntry(moduleUrl, argvPath) {
	if (!argvPath) {
		return false;
	}
	const canonical = (path) => {
		try {
			return realpathSync(path);
		} catch {
			return resolve(path);
		}
	};
	return canonical(fileURLToPath(moduleUrl)) === canonical(argvPath);
}

if (isCliEntry(import.meta.url, process.argv[1])) {
	const distDir = process.argv.find((arg, i) => i > 1 && !arg.startsWith('--'));
	const { problems, sizes, kept } = await checkTreeShaking(distDir && resolve(distDir));
	if (process.argv.includes('--json')) {
		console.log(JSON.stringify(sizes, null, '\t'));
	} else {
		for (const [name, { bytes, budget }] of Object.entries(sizes)) {
			console.log(`${name}: ${bytes} / ${budget} bytes`);
		}
		console.log(`side-effect import keeps: ${kept.join(', ') || 'nothing'}`);
	}
	if (problems.length) {
		console.error(problems.map((problem) => `  - ${problem}`).join('\n'));
		process.exit(1);
	}
}
