#!/usr/bin/env node
/**
 * check-shipped-css.mjs: fail when a package publishes CSS that a consumer's bundler cannot minify.
 *
 * Every `.css` file a package would publish is minified with esbuild and with Lightning CSS, the
 * two minifiers consumers' bundlers use (Vite 8 minifies CSS with Lightning CSS, esbuild-based
 * tools with esbuild). Any warning or error from either fails the check: a malformed stylesheet
 * minifies "successfully" with warnings while the bundler silently drops the broken rules.
 *
 *   node scripts/check-shipped-css.mjs [<dir prefix>...]
 *
 * A prefix (relative to the repository root, for example `viewers/pptx` or `src/ui`) selects the
 * publishable packages under it; none selects every package. Run it after the packages are built.
 * A package may publish no CSS file (ooxml-ui and the Word, Excel, Visio and Teams packages keep
 * their styles inside their modules), but a `.css` file its `exports` name must be in the package,
 * so an unbuilt package fails instead of passing with nothing checked.
 */

import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';

import { transform as esbuildTransform } from 'esbuild';
import { transform as lightningTransform } from 'lightningcss';

import { PACKAGES, presentPackages } from './release-plan.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

/**
 * Tailwind directives, which Lightning CSS reports as unknown. A stylesheet that uses them (the
 * React package's `theme.css`) is published for the host's own Tailwind build to import, which
 * compiles them away before any minifier sees them; esbuild still parses it and must not warn.
 */
const TAILWIND_AT_RULE =
	/^Unknown at rule: @(theme|source|utility|variant|custom-variant|apply|config|plugin|reference)\b/u;

/** Every problem esbuild and Lightning CSS report when minifying `css`, one line each. */
export async function cssProblems(css, filename) {
	const problems = [];
	try {
		const { warnings } = await esbuildTransform(css, {
			loader: 'css',
			minify: true,
			sourcefile: filename,
			logLevel: 'silent',
		});
		for (const warning of warnings) problems.push(`esbuild: ${esbuildMessage(warning)}`);
	} catch (error) {
		for (const message of error.errors ?? [error])
			problems.push(`esbuild: ${esbuildMessage(message)}`);
	}
	try {
		const { warnings } = lightningTransform({
			filename,
			code: Buffer.from(css),
			minify: true,
			errorRecovery: true,
		});
		for (const warning of warnings.filter((w) => !TAILWIND_AT_RULE.test(w.message)))
			problems.push(`lightningcss: ${lightningMessage(warning)}`);
	} catch (error) {
		problems.push(`lightningcss: ${lightningMessage(error)}`);
	}
	return problems;
}

const esbuildMessage = ({ text, location }) =>
	location ? `${text} at ${location.line}:${location.column}` : String(text);

const lightningMessage = ({ message, loc }) =>
	loc ? `${message} at ${loc.line}:${loc.column}` : String(message);

/** The publishable packages under `prefixes` (all with none), with the directory they are packed from. */
export function packagesUnder(prefixes, root = ROOT) {
	return Object.values(presentPackages(root, PACKAGES)).flatMap((meta) => {
		const manifest = JSON.parse(readFileSync(join(root, meta.dir, 'package.json'), 'utf8'));
		const selected =
			prefixes.length === 0 ||
			prefixes.some((prefix) => meta.dir === prefix || meta.dir.startsWith(`${prefix}/`));
		if (manifest.private || !selected) return [];
		return [{ npm: meta.npm, dir: meta.dir, packDir: join(root, meta.dir, meta.publishDir ?? '') }];
	});
}

/** The `.css` files a manifest's `exports` name, relative to the package, without `./`. */
export function exportedCss(manifest) {
	const targets = [];
	const walk = (value) => {
		if (typeof value === 'string') {
			if (value.endsWith('.css')) targets.push(value.replace(/^\.\//u, ''));
		} else if (value && typeof value === 'object') {
			Object.values(value).forEach(walk);
		}
	};
	walk(manifest.exports);
	return [...new Set(targets)];
}

/**
 * The files packing `packDir` would publish, relative to it and `/`-separated: what the manifest's
 * `files` entries name (a file or a directory, whole), or the whole directory without `files`.
 * `npm pack --dry-run` gives the exact list but takes seconds per package; no package here uses
 * globs in `files` or an `.npmignore` that drops CSS.
 */
export function packedFiles(packDir, manifest) {
	const roots = Array.isArray(manifest.files)
		? manifest.files.map((entry) => entry.replace(/^\.\//u, '').replace(/\/+$/u, ''))
		: [''];
	const out = [];
	const walk = (path) => {
		const absolute = join(packDir, path);
		if (!existsSync(absolute)) return;
		if (!statSync(absolute).isDirectory()) {
			out.push(path);
			return;
		}
		for (const name of readdirSync(absolute)) {
			if (name !== 'node_modules' && name !== '.git') walk(path ? `${path}/${name}` : name);
		}
	};
	roots.forEach(walk);
	return out;
}

async function main(prefixes) {
	let failed = false;
	let checked = 0;
	for (const pkg of packagesUnder(prefixes)) {
		const manifest = JSON.parse(readFileSync(join(pkg.packDir, 'package.json'), 'utf8'));
		const files = packedFiles(pkg.packDir, manifest);
		for (const missing of exportedCss(manifest).filter((path) => !files.includes(path))) {
			console.error(`${pkg.npm} exports ${missing}, which is not in the package: build it first.`);
			failed = true;
		}
		for (const path of files.filter((file) => file.endsWith('.css'))) {
			checked += 1;
			const file = join(pkg.packDir, path);
			const problems = await cssProblems(readFileSync(file, 'utf8'), relative(ROOT, file));
			for (const problem of problems) console.error(`${pkg.npm} ${path}: ${problem}`);
			failed ||= problems.length > 0;
		}
	}
	console.log(
		`Minified ${checked} published CSS files under ${prefixes.join(', ') || 'the repository'}.`,
	);
	return failed ? 1 : 0;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
	process.exitCode = await main(process.argv.slice(2));
}
