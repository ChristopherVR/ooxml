#!/usr/bin/env node
/**
 * check-lockfile-peers.mjs: fail when `bun.lock` has a peer dependency that Bun's isolated linker
 * links to a package the lockfile cannot name for it.
 *
 * Bun 1.4.4 and 1.4.5 abort a cold `bun install --frozen-lockfile` (empty cache) with
 * `panic: index out of bounds: ... the index is 4294967295` for this shape. A peer's lockfile slot
 * is found by looking up its name from the dependent's key outwards (`a/b/<peer>`, `a/<peer>`, then
 * the root's `<peer>`). When that lookup finds nothing, the slot is empty, but the isolated linker
 * still links the peer to the nearest package of that name, which the lockfile keeps nested under
 * some workspace. If the linker reaches that package through the peer first, it queues the tarball
 * download from the peer's dependency, and on completion reads the package from the empty slot.
 * A warm cache downloads nothing, which is why only cold installs crash.
 *
 * The repository hit it with `@analogjs/vite-plugin-angular` (optional peer `@angular/build`, which
 * the lock nested under every Angular workspace because the root's TypeScript did not satisfy its
 * peer). The reproduction is recorded with the 2.0 CI evidence. Until Bun reads the package from the
 * download task, keep this check green: hoist the package (give the root a version its peers accept)
 * rather than leave the peer's slot empty.
 *
 *   node scripts/check-lockfile-peers.mjs [--file bun.lock]
 */

import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseLock } from './check-lockfile-drift.mjs';

/** A lockfile key's package names, outermost first: `a/@s/b/c` is `a`, `@s/b`, `c`. */
export function keySegments(key) {
	const parts = key.split('/');
	const out = [];
	for (let i = 0; i < parts.length; i++)
		out.push(parts[i].startsWith('@') ? `${parts[i]}/${parts[++i]}` : parts[i]);
	return out;
}

/**
 * Peers whose lockfile lookup finds no package while the lockfile nests a package of that name
 * elsewhere, as `{ key, peer, nested }` (`nested`: the keys the linker could link instead).
 */
export function emptyPeerSlots(lock) {
	const packages = lock.packages ?? {};
	const nestedByName = new Map();
	for (const key of Object.keys(packages)) {
		const segments = keySegments(key);
		if (segments.length < 2) continue;
		const name = segments.at(-1);
		(nestedByName.get(name) ?? nestedByName.set(name, []).get(name)).push(key);
	}
	const found = [];
	for (const [key, entry] of Object.entries(packages)) {
		const meta = entry.find((part) => part && typeof part === 'object' && !Array.isArray(part));
		const segments = keySegments(key);
		for (const peer of Object.keys(meta?.peerDependencies ?? {})) {
			const nested = nestedByName.get(peer);
			if (!nested) continue;
			let resolved = false;
			for (let depth = segments.length; depth >= 0 && !resolved; depth--)
				resolved = [...segments.slice(0, depth), peer].join('/') in packages;
			if (!resolved) found.push({ key, peer, nested });
		}
	}
	return found;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
	const args = process.argv.slice(2);
	const file = args.includes('--file') ? args[args.indexOf('--file') + 1] : 'bun.lock';
	const found = emptyPeerSlots(parseLock(readFileSync(file, 'utf8')));
	for (const { key, peer, nested } of found)
		console.error(
			`${file}: the peer ${peer} of ${key} has no lockfile slot, but is nested at ${nested.join(', ')}`,
		);
	if (found.length > 0) {
		console.error(
			'A cold `bun install` aborts on this shape (scripts/check-lockfile-peers.mjs explains why).',
		);
		process.exitCode = 1;
	} else console.log(`${file}: every peer the lockfile nests has a slot.`);
}
