// File storage bookkeeping of the reference server: which scope (a workspace id, or a private chat
// room name) a file belongs to, who uploaded it, how many bytes each workspace uses, and the sweep
// of partial uploads. The bytes themselves are streamed by index.mjs.
import { existsSync, mkdirSync, readdirSync, readFileSync, statSync, unlinkSync } from 'node:fs';
import { join } from 'node:path';
import { isChatRoom, shortHash } from './chat-access.mjs';
import { writeFileAtomic } from './room-store.mjs';

const PART = /^\..+\.part$/u;
const CHAT_HASH = /^chat-([0-9a-f]{16})-/u;

/** Scope ids are a workspace id or a chat room name; both match this. */
export const SCOPE = /^[\w-]{1,128}$/u;

/** The quota bucket of a scope: the workspace hash (chat rooms carry it in their name). */
export const quotaKeyOf = (scope) =>
	isChatRoom(scope) ? CHAT_HASH.exec(scope)[1] : shortHash(scope);

const quiet = (fn, fallback) => {
	try {
		return fn();
	} catch {
		return fallback;
	}
};

export function createFileStore({ filesDir, metaDir, log = () => {} }) {
	mkdirSync(filesDir, { recursive: true });
	mkdirSync(metaDir, { recursive: true });
	/** @type {Map<string, number>} quota key -> bytes stored */
	const usage = new Map();
	const dirOf = (scope) => join(filesDir, scope);
	const metaPath = (scope) => join(metaDir, `${scope}.json`);

	const readOwners = (scope) =>
		quiet(() => {
			const parsed = JSON.parse(readFileSync(metaPath(scope), 'utf8'));
			return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
		}, {});
	const writeOwners = (scope, owners) => {
		try {
			writeFileAtomic(metaPath(scope), Buffer.from(JSON.stringify(owners)));
		} catch (error) {
			log(`files ${scope}: could not persist uploader records (${error.message})`);
		}
	};
	const scopes = () =>
		quiet(
			() => readdirSync(filesDir, { withFileTypes: true }).filter((e) => e.isDirectory()),
			[],
		).map((e) => e.name);
	const filesOf = (scope) =>
		quiet(
			() =>
				readdirSync(dirOf(scope), { withFileTypes: true })
					.filter((e) => e.isFile() && !PART.test(e.name))
					.map((e) => e.name),
			[],
		);
	const sizeOf = (scope, name) => quiet(() => statSync(join(dirOf(scope), name)).size, 0);

	for (const scope of scopes()) {
		const key = quotaKeyOf(scope);
		for (const name of filesOf(scope)) usage.set(key, (usage.get(key) ?? 0) + sizeOf(scope, name));
	}

	return {
		dirOf,
		bytesOf: (scope) => usage.get(quotaKeyOf(scope)) ?? 0,
		totalBytes: () => [...usage.values()].reduce((a, b) => a + b, 0),
		/** Record a stored file and who uploaded it. */
		added(scope, name, size, user) {
			const key = quotaKeyOf(scope);
			usage.set(key, (usage.get(key) ?? 0) + size);
			if (user) writeOwners(scope, { ...readOwners(scope), [name]: user });
		},
		uploaderOf: (scope, name) => readOwners(scope)[name] ?? null,
		list(scope) {
			const owners = readOwners(scope);
			return filesOf(scope)
				.map((name) => {
					const stat = quiet(() => statSync(join(dirOf(scope), name)), null);
					return {
						name,
						size: stat?.size ?? 0,
						modified: stat ? stat.mtime.toISOString() : null,
						uploader: owners[name] ?? null,
					};
				})
				.sort((a, b) => (a.name < b.name ? -1 : 1));
		},
		/** Delete one file; false when it does not exist. */
		remove(scope, name) {
			const file = join(dirOf(scope), name);
			if (!existsSync(file)) return false;
			const size = sizeOf(scope, name);
			unlinkSync(file);
			const key = quotaKeyOf(scope);
			usage.set(key, Math.max(0, (usage.get(key) ?? 0) - size));
			const owners = readOwners(scope);
			if (name in owners) {
				delete owners[name];
				writeOwners(scope, owners);
			}
			return true;
		},
		/**
		 * Remove partial uploads older than `maxAgeMs` (0: all of them, for a start-up sweep when
		 * nothing is in flight) and uploader records of files that are gone. Returns the count.
		 */
		sweep(maxAgeMs = 0, now = Date.now()) {
			let removed = 0;
			for (const scope of scopes()) {
				const dir = dirOf(scope);
				for (const entry of quiet(() => readdirSync(dir, { withFileTypes: true }), [])) {
					if (!entry.isFile() || !PART.test(entry.name)) continue;
					const age = now - quiet(() => statSync(join(dir, entry.name)).mtimeMs, now);
					if (maxAgeMs > 0 && age < maxAgeMs) continue;
					if (quiet(() => (unlinkSync(join(dir, entry.name)), true), false)) removed++;
				}
				const owners = readOwners(scope);
				const present = new Set(filesOf(scope));
				const stale = Object.keys(owners).filter((name) => !present.has(name));
				if (stale.length) {
					for (const name of stale) delete owners[name];
					writeOwners(scope, owners);
				}
			}
			if (removed) log(`files: removed ${removed} partial upload(s)`);
			return removed;
		},
	};
}
