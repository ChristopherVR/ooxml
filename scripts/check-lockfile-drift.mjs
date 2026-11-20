#!/usr/bin/env node
/**
 * check-lockfile-drift.mjs: fail when `bun.lock` changed in any way other than the `version` of a
 * workspace package.
 *
 * The release job bumps package versions and then runs a plain `bun install`, because the lockfile
 * records every workspace's version and `--frozen-lockfile` would refuse the bump. That install
 * could also resolve a newer third-party package and the release commit would ship it unreviewed.
 * This check compares the lockfile on disk with the one at a git revision (default `HEAD`) and
 * allows only `workspaces.<path>.version` to differ, plus a workspace's ranges on other workspace
 * packages (the planner raises a released viewer's `ooxml-core` range to the version it releases),
 * which resolve inside the repository and so cannot move a third-party resolution.
 *
 * Those versions and ranges must then equal the manifests on disk. The release job commits its own
 * manifests next to a lockfile the prepare job wrote, so a lockfile that records another version
 * (an install that did not rewrite it, or a plan that moved in between) would land on `main` and
 * fail the next `bun install --frozen-lockfile`.
 *
 *   node scripts/check-lockfile-drift.mjs [--base <rev>] [--file bun.lock]
 */

import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

/** Parses a `bun.lock` (JSON with trailing commas). */
export function parseLock(text) {
	return JSON.parse(text.replace(/,(\s*[}\]])/gu, '$1'));
}

const DEPENDENCY_FIELDS = [
	'dependencies',
	'devDependencies',
	'optionalDependencies',
	'peerDependencies',
];

/**
 * A copy of the lockfile without what the release job bumps: each workspace package's version and
 * each workspace's ranges on other workspace packages.
 */
export function withoutWorkspaceVersions(lock) {
	const copy = structuredClone(lock);
	const names = siblingNames(copy);
	for (const workspace of Object.values(copy.workspaces ?? {})) {
		delete workspace.version;
		for (const field of DEPENDENCY_FIELDS) {
			const ranges = workspace[field];
			if (!ranges) continue;
			for (const name of Object.keys(ranges)) if (names.has(name)) delete ranges[name];
		}
	}
	return copy;
}

const siblingNames = (lock) =>
	new Set(
		Object.values(lock.workspaces ?? {})
			.map((workspace) => workspace.name)
			.filter(Boolean),
	);

/**
 * Lockfile paths whose workspace version, or range on another workspace package, is not what the
 * workspace's manifest says. `readManifest(path)` returns the parsed `package.json` of the
 * workspace at `path` (`''` is the root), or null when there is none.
 */
export function workspaceMismatches(lock, readManifest) {
	const names = siblingNames(lock);
	const out = [];
	for (const [path, workspace] of Object.entries(lock.workspaces ?? {})) {
		const at = `bun.lock.workspaces.${path}`;
		const manifest = readManifest(path);
		if (!manifest) {
			out.push(`${at} (no package.json)`);
			continue;
		}
		if (workspace.version !== manifest.version) out.push(`${at}.version`);
		for (const field of DEPENDENCY_FIELDS) {
			for (const name of names) {
				if (workspace[field]?.[name] !== manifest[field]?.[name])
					out.push(`${at}.${field}.${name}`);
			}
		}
	}
	return out;
}

/**
 * Lockfile paths that differ, ignoring workspace versions. Empty when only versions changed. With
 * `readManifest` (see `workspaceMismatches`), the workspace versions and sibling ranges of the
 * current lockfile must also equal the manifests.
 */
export function lockDrift(baseText, currentText, readManifest) {
	const lock = parseLock(currentText);
	const base = withoutWorkspaceVersions(parseLock(baseText));
	const current = withoutWorkspaceVersions(lock);
	const drift = readManifest ? workspaceMismatches(lock, readManifest) : [];
	const walk = (a, b, path) => {
		if (JSON.stringify(a) === JSON.stringify(b)) return;
		const isObject = (value) => value && typeof value === 'object' && !Array.isArray(value);
		if (isObject(a) && isObject(b)) {
			for (const key of new Set([...Object.keys(a), ...Object.keys(b)])) {
				walk(a[key], b[key], `${path}.${key}`);
			}
		} else {
			drift.push(path);
		}
	};
	walk(base, current, 'bun.lock');
	return drift;
}

function main() {
	const argv = process.argv.slice(2);
	const value = (flag, fallback) => (argv.includes(flag) ? argv[argv.indexOf(flag) + 1] : fallback);
	const file = value('--file', 'bun.lock');
	const base = value('--base', 'HEAD');
	const baseText = execFileSync('git', ['show', `${base}:${file}`], {
		encoding: 'utf8',
		maxBuffer: 64 * 1024 * 1024,
	});
	const root = dirname(file);
	const readManifest = (path) => {
		const manifest = join(root, path, 'package.json');
		return existsSync(manifest) ? JSON.parse(readFileSync(manifest, 'utf8')) : null;
	};
	const drift = lockDrift(baseText, readFileSync(file, 'utf8'), readManifest);
	if (drift.length === 0) {
		console.log(
			`${file} differs from ${base} only in workspace package versions (if at all), and they match the manifests.`,
		);
		return;
	}
	console.error(
		`::error::${file} changed beyond workspace versions, or does not match the manifests:`,
	);
	for (const path of drift.slice(0, 50)) console.error(`  ${path}`);
	if (drift.length > 50) console.error(`  ... and ${drift.length - 50} more`);
	process.exit(1);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
