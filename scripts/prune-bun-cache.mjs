#!/usr/bin/env node
// Removes from Bun's install cache every package version the lockfile no longer resolves.
//
//   node scripts/prune-bun-cache.mjs <cache-dir> <bun.lock>
//
// CI restores the cache from the newest earlier key when the lockfile's packages change, installs,
// and saves the result under the new key, so without pruning the cache keeps every version it has
// ever held (842 MB against about 330 MB for one lockfile, in November 2026), and every job pays to
// extract it. Only `<name>@<version>@@@<n>` folders and the links to them are touched; anything
// else in the cache is left as it is.
import { lstatSync, readFileSync, readdirSync, readlinkSync, rmSync } from 'node:fs';
import { basename, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseLock } from './ci-manifests.mjs';

const VERSIONED = /^(.+)@([^@/]+)@@@\d+$/;

/** The `name@version` of every package the lockfile resolves. */
export function lockedPackages(lockText) {
	const locked = new Set();
	for (const entry of Object.values(parseLock(lockText).packages ?? {}))
		if (Array.isArray(entry) && typeof entry[0] === 'string') locked.add(entry[0]);
	return locked;
}

/** The cache paths (relative, `/`-separated) of versions `locked` does not name. */
export function staleEntries(paths, locked) {
	return paths.filter((path) => {
		const match = VERSIONED.exec(path);
		return match !== null && !locked.has(`${match[1]}@${match[2]}`);
	});
}

/** Every versioned folder of the cache, as `name@version@@@n` or `@scope/name@version@@@n`. */
function versionedFolders(cache) {
	const out = [];
	for (const name of readdirSync(cache)) {
		if (name.startsWith('@')) {
			for (const inner of readdirSync(join(cache, name)))
				if (VERSIONED.test(inner)) out.push(`${name}/${inner}`);
		} else if (VERSIONED.test(name)) out.push(name);
	}
	return out;
}

/** Deletes the stale versions and the per-package links to them. Returns how many were removed. */
export function pruneCache(cache, locked) {
	const stale = staleEntries(versionedFolders(cache), locked);
	const removed = new Set(stale.map((path) => resolve(cache, path)));
	for (const path of removed) rmSync(path, { recursive: true, force: true });
	// `<name>/<version>@@@n` links point at the versioned folders; drop the ones left dangling.
	const linkDirs = readdirSync(cache).flatMap((name) =>
		name.startsWith('@')
			? readdirSync(join(cache, name)).map((inner) => join(cache, name, inner))
			: [join(cache, name)],
	);
	for (const dir of linkDirs) {
		if (!lstatSync(dir).isDirectory() || VERSIONED.test(basename(dir))) continue;
		for (const link of readdirSync(dir)) {
			const path = join(dir, link);
			if (lstatSync(path).isSymbolicLink() && removed.has(resolve(dir, readlinkSync(path))))
				rmSync(path, { force: true });
		}
	}
	return stale.length;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
	const [cache, lock] = process.argv.slice(2);
	if (!cache || !lock) throw new Error('usage: prune-bun-cache.mjs <cache-dir> <bun.lock>');
	const count = pruneCache(cache, lockedPackages(readFileSync(lock, 'utf8')));
	console.log(`Removed ${count} package versions bun.lock no longer resolves from ${cache}.`);
}
