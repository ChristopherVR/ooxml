/**
 * Bundles the browser entry points of installed packages the way a web host does (esbuild,
 * `platform: 'browser'`, no fallbacks) and reports every Node built-in or Node-only package they
 * can reach, and every dynamic `import()` whose specifier the bundler cannot see. The package
 * checks run it on the packed packages: `src/core/scripts/package-smoke.mjs` (ooxml-core),
 * `src/ui/scripts/package-smoke.mjs` (ooxml-core and ooxml-ui) and
 * `viewers/pptx/scripts/check-binding-packages.mjs` (the PowerPoint binding packages).
 */
import { existsSync, readdirSync, readFileSync, realpathSync } from 'node:fs';
import { builtinModules } from 'node:module';
import { join, relative, sep } from 'node:path';

import * as esbuild from 'esbuild';

/** Packages that only run under Node.js (native addons and server DOMs). */
export const NODE_ONLY_PACKAGES = ['@napi-rs/canvas', 'canvas', 'jsdom', 'sharp'];

/** Public subpaths that are Node-only by design and so are not browser entry points. */
const NODE_SUBPATH = /(?:^|\/)(?:cli|signature-node|node)$/;
/** Files a browser bundle cannot load: styles, fonts and images are the host's concern. */
const ASSET = /\.(?:css|less|scss|woff2?|ttf|otf|eot|png|jpe?g|gif|svg|webp)$/;
/** Shared chunk files a wildcard export happens to expose; they are reached through the entries. */
const CHUNK_FILE = /(?:^|\/)chunk-[^/]*$/;
const HASHED_FILE = /-[A-Z0-9]{8}$/;

const builtins = new Set(builtinModules.flatMap((name) => [name, `node:${name}`]));

/** Whether an import specifier names a Node built-in or a Node-only package. */
export function isNodeOnlySpecifier(specifier) {
	if (specifier.startsWith('node:') || builtins.has(specifier)) {
		return true;
	}
	if (builtins.has(specifier.split('/')[0])) {
		return true;
	}
	return NODE_ONLY_PACKAGES.some((name) => specifier === name || specifier.startsWith(`${name}/`));
}

function packageDir(root, name) {
	return join(root, 'node_modules', ...name.split('/'));
}

function runtimeTarget(value) {
	if (typeof value === 'string') {
		return value;
	}
	if (!value || typeof value !== 'object') {
		return undefined;
	}
	for (const key of ['browser', 'import', 'module', 'default']) {
		const target = runtimeTarget(value[key]);
		if (target) {
			return target;
		}
	}
	return undefined;
}

/** The files a `./x/*` export matches, as subpaths (`./x/a`, `./x/b/c`). */
function wildcardSubpaths(dir, subpath, target) {
	const [prefix, suffix] = target.split('*');
	const base = join(dir, prefix);
	if (!prefix.endsWith('/') || !existsSync(base)) {
		return [];
	}
	return readdirSync(base, { recursive: true })
		.map((file) => String(file).split(sep).join('/'))
		.filter((file) => file.endsWith(suffix))
		.map((file) => file.slice(0, file.length - suffix.length))
		.filter((stem) => !stem.endsWith('.d') && !CHUNK_FILE.test(stem) && !HASHED_FILE.test(stem))
		.map((stem) => subpath.replaceAll('*', stem));
}

/** Every public browser entry point of an installed package, as import specifiers. */
export function browserEntries(root, name) {
	const dir = packageDir(root, name);
	const manifest = JSON.parse(readFileSync(join(dir, 'package.json'), 'utf8'));
	const exports = manifest.exports ?? { '.': manifest.module ?? manifest.main };
	const entries = [];
	for (const [subpath, value] of Object.entries(exports)) {
		const target = runtimeTarget(value);
		if (!target || ASSET.test(target) || target.endsWith('.json') || subpath.endsWith('.json')) {
			continue;
		}
		const subpaths = subpath.includes('*') ? wildcardSubpaths(dir, subpath, target) : [subpath];
		for (const each of subpaths) {
			const specifier = each === '.' ? name : `${name}/${each.slice(2)}`;
			if (!NODE_SUBPATH.test(specifier)) {
				entries.push(specifier);
			}
		}
	}
	return [...new Set(entries)];
}

/**
 * Bundle each entry for the browser and return the Node-only modules it reaches.
 *
 * @param {object} options
 * @param {string} options.root Directory whose `node_modules` holds the packages.
 * @param {string[]} options.packages Packages whose browser entries are checked.
 * @returns {Promise<{ entries: string[], findings: Array<{ entry: string, reason: string }> }>}
 *   The entries bundled, and one finding per entry and problem.
 */
export async function findNodeOnlyModules({ root: rootDir, packages }) {
	// esbuild reports real paths; compare them with real paths (macOS /tmp is a symlink).
	const root = realpathSync(rootDir);
	const entries = packages.flatMap((name) => browserEntries(root, name));
	const imports = [];
	const plugin = {
		name: 'browser-bundle-guard',
		setup(build) {
			build.onResolve({ filter: /^guard-entry:/ }, (args) => ({
				path: args.path.slice('guard-entry:'.length),
				namespace: 'guard-entry',
			}));
			build.onLoad({ filter: /.*/, namespace: 'guard-entry' }, (args) => ({
				contents: `export * from ${JSON.stringify(args.path)};`,
				resolveDir: root,
				loader: 'js',
			}));
			build.onResolve({ filter: /.*/ }, async (args) => {
				if (args.pluginData?.guard || args.namespace === 'guard-entry') {
					return undefined;
				}
				if (isNodeOnlySpecifier(args.path)) {
					imports.push({ specifier: args.path, importer: args.importer, kind: args.kind });
					return { path: args.path, external: true };
				}
				if (/^[./]/.test(args.path) || args.path.startsWith('#')) {
					return undefined;
				}
				// A package the host has not installed (an optional peer such as the AI SDK) is
				// the host's choice, not a Node-only module: leave it to the host's bundler.
				const resolved = await build.resolve(args.path, {
					kind: args.kind,
					importer: args.importer,
					resolveDir: args.resolveDir,
					pluginData: { guard: true },
				});
				return resolved.errors.length > 0 ? { path: args.path, external: true } : undefined;
			});
		},
	};
	const result = await esbuild.build({
		entryPoints: entries.map((entry) => ({ in: `guard-entry:${entry}`, out: entry })),
		absWorkingDir: root,
		outdir: join(root, '.browser-bundle-guard'),
		bundle: true,
		splitting: true,
		format: 'esm',
		platform: 'browser',
		write: false,
		metafile: true,
		logLevel: 'silent',
		// A dynamic `import()` whose specifier is not a string literal: no bundler can see what it
		// loads, so it can hide a Node-only package from this check (and a minifier may fold the
		// specifier back into a literal that the host's bundler then fails to resolve). esbuild
		// reports each one, also inside node_modules, once its log level is raised.
		logOverride: { 'unsupported-dynamic-import': 'warning' },
		loader: Object.fromEntries(
			['.css', '.less', '.scss', '.woff', '.woff2', '.ttf', '.otf', '.eot', '.png', '.jpg']
				.concat(['.jpeg', '.gif', '.svg', '.webp'])
				.map((extension) => [extension, 'empty']),
		),
		plugins: [plugin],
	});

	// Which entries each output file belongs to, following static and dynamic chunk imports.
	const outputs = result.metafile.outputs;
	const computed = result.warnings
		.filter((warning) => warning.id === 'unsupported-dynamic-import' && warning.location)
		.map(({ location }) => ({ file: location.file, line: location.line }));
	const outputKey = (path) => relative(root, path).split(sep).join('/');
	const reach = (start) => {
		const seen = new Set();
		const visit = (key) => {
			if (seen.has(key) || !outputs[key]) {
				return;
			}
			seen.add(key);
			for (const { path } of outputs[key].imports) {
				visit(path);
			}
		};
		visit(start);
		return seen;
	};
	const findings = [];
	for (const [key, output] of Object.entries(outputs)) {
		// Lazily imported modules are outputs with an entry point too; they are reached below.
		if (!output.entryPoint?.startsWith('guard-entry:')) {
			continue;
		}
		const entry = output.entryPoint.replace(/^guard-entry:/, '');
		const inputs = new Set([...reach(key)].flatMap((each) => Object.keys(outputs[each].inputs)));
		for (const { specifier, importer, kind } of imports) {
			if (inputs.has(outputKey(importer))) {
				findings.push({ entry, reason: `${kind} of ${specifier} from ${outputKey(importer)}` });
			}
		}
		for (const { file, line } of computed) {
			if (inputs.has(file)) {
				findings.push({
					entry,
					reason: `dynamic import with a computed specifier at ${file}:${line}`,
				});
			}
		}
	}
	return { entries, findings: dedupe(findings) };
}

/** Throw when any finding is reported, listing them; otherwise return the entry count. */
export async function assertNoNodeOnlyModules(options) {
	const { entries, findings } = await findNodeOnlyModules(options);
	if (findings.length > 0) {
		const lines = findings.map(({ entry, reason }) => `  ${entry}: ${reason}`);
		throw new Error(`Browser entries reach Node-only modules:\n${lines.join('\n')}`);
	}
	return entries.length;
}

function dedupe(findings) {
	return [
		...new Map(findings.map((finding) => [`${finding.entry}|${finding.reason}`, finding])).values(),
	];
}
