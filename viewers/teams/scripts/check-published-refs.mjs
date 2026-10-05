/**
 * Guard: nothing a published package ships may import a module its manifest does not declare.
 *
 * Ported from ChristopherVR/docx-viewer `scripts/check-published-refs.mjs`. The framework packages
 * inline the private web component (`teams-viewer`); the logic is imported from ooxml-core and the
 * visual primitives from ooxml-ui, both real registry dependencies. If bundling or declaration
 * flattening misses something, the tarball imports an unpublished package a consumer cannot
 * install.
 *
 *   node scripts/check-published-refs.mjs [react vue angular svelte solid vanilla server]
 *
 * It scans every `.js`, `.mjs`, `.d.ts` and `.svelte` file `npm pack` would include, so the
 * framework packages must be built first (`bun run build:packages`).
 */
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { builtinModules } from 'node:module';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');

/** Short name -> directory of every published package. */
export const PUBLISHED_DIRS = {
	react: 'packages/react',
	vue: 'packages/vue',
	angular: 'packages/angular',
	svelte: 'packages/svelte',
	solid: 'packages/solid',
	vanilla: 'packages/vanilla',
	server: 'server',
};

// `from '<x>'`, bare `import '<x>'`, `import('<x>')` and `require('<x>')`; the specifier must look
// like a module id so prose such as `from": "` inside a string never matches.
const SPECIFIER =
	/(?:\bfrom\s*|\bimport\s*\(?\s*|\brequire\s*\(\s*)["']((?:@[\w.-]+\/)?[\w.-]+(?:\/[\w./-]*)?|\.{1,2}\/[\w./-]*)["']/gu;

/** Module specifiers a source file imports, de-duplicated, relative ones excluded. */
export function importedPackages(source) {
	const found = new Set();
	for (const [, specifier] of source.matchAll(SPECIFIER)) {
		if (!specifier.startsWith('.')) found.add(specifier);
	}
	return [...found];
}

const packageOf = (specifier) =>
	specifier.startsWith('@') ? specifier.split('/').slice(0, 2).join('/') : specifier.split('/')[0];

/** Specifiers in `source` that are neither Node built-ins nor declared by `manifest`. */
export function undeclaredImports(source, manifest) {
	const declared = new Set([
		...Object.keys(manifest.dependencies ?? {}),
		...Object.keys(manifest.peerDependencies ?? {}),
	]);
	return importedPackages(source).filter((specifier) => {
		const name = packageOf(specifier);
		return !(name.startsWith('node:') || builtinModules.includes(name) || declared.has(name));
	});
}

/**
 * Names a published manifest may never depend on: the private workspace packages (`teams-viewer`,
 * the `teams-demo-*` apps and the pre-rename `teams-*` names), anything under `@christophervr/`
 * (ole2 is inlined inside ooxml-core) and any `ooxml-*` package other than core and ui.
 */
export const INTERNAL_NAME = /^(?:@christophervr\/|teams-|ooxml-(?!core$|ui$))/u;

/** Manifest entries a consumer cannot install: internal packages and local protocols. */
export function forbiddenManifestEntries(manifest) {
	const found = [];
	for (const field of ['dependencies', 'peerDependencies', 'optionalDependencies']) {
		for (const [name, range] of Object.entries(manifest[field] ?? {})) {
			if (INTERNAL_NAME.test(name) || /^(?:workspace|file|link):/u.test(range))
				found.push(`${field}.${name}`);
		}
	}
	return found;
}

function packedFiles(dir) {
	const output = execFileSync(
		'npm',
		['pack', '--dry-run', '--ignore-scripts', '--json', '--workspaces=false'],
		{
			cwd: dir,
			encoding: 'utf8',
			stdio: ['ignore', 'pipe', 'pipe'],
			shell: process.platform === 'win32',
		},
	);
	return JSON.parse(output)[0].files.map((file) => file.path);
}

/** Scan one packed package; returns the offending files and manifest entries. */
export function checkPackage(name) {
	const dir = join(ROOT, PUBLISHED_DIRS[name]);
	const manifest = JSON.parse(readFileSync(join(dir, 'package.json'), 'utf8'));
	const files = packedFiles(dir).filter((path) => /\.(?:m?js|d\.ts|svelte)$/u.test(path));
	if (files.length === 0)
		throw new Error(`${PUBLISHED_DIRS[name]} has no built files; run build:packages`);
	const offenders = [];
	for (const file of files) {
		const bad = undeclaredImports(readFileSync(join(dir, file), 'utf8'), manifest);
		if (bad.length > 0) offenders.push({ file, specifiers: bad });
	}
	return {
		name: manifest.name,
		scanned: files.length,
		offenders,
		manifest: forbiddenManifestEntries(manifest),
	};
}

function main() {
	const requested = process.argv.slice(2);
	const names = requested.length > 0 ? requested : Object.keys(PUBLISHED_DIRS);
	let failed = false;
	for (const name of names) {
		if (!(name in PUBLISHED_DIRS)) throw new Error(`Unknown package "${name}".`);
		const result = checkPackage(name);
		if (result.offenders.length === 0 && result.manifest.length === 0) {
			console.log(
				`[check-published-refs] ${result.name}: ${result.scanned} files, no undeclared imports.`,
			);
			continue;
		}
		failed = true;
		for (const { file, specifiers } of result.offenders)
			console.error(
				`[check-published-refs] ${result.name}: ${file} imports ${specifiers.join(', ')}`,
			);
		for (const entry of result.manifest)
			console.error(`[check-published-refs] ${result.name}: manifest ${entry} cannot be installed`);
	}
	if (failed) process.exitCode = 1;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
