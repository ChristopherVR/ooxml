// Room snapshot storage: atomic writes, compaction and the sweep of leftovers. A snapshot is one
// Yjs update per room. It is written to a temporary file, flushed to disk and renamed over the old
// one, so a crash at any point leaves either the old snapshot or the new one, never a torn file.
import {
	closeSync,
	fsyncSync,
	openSync,
	readdirSync,
	readFileSync,
	renameSync,
	statSync,
	unlinkSync,
	writeSync,
} from 'node:fs';
import { randomBytes } from 'node:crypto';
import { join } from 'node:path';
import * as Y from 'yjs';

const realFs = { openSync, writeSync, fsyncSync, closeSync, renameSync, unlinkSync };
const TEMP = /\.[0-9a-f]{8}\.tmp$/u;

/**
 * Write `bytes` to `file` atomically (temp file, fsync, rename). `fs` can replace the file
 * operations, which is how the tests cut the write short. On any failure the temp file is removed
 * and the previous content of `file` is untouched.
 */
export function writeFileAtomic(file, bytes, fs = realFs) {
	const tmp = `${file}.${randomBytes(4).toString('hex')}.tmp`;
	let fd;
	try {
		fd = fs.openSync(tmp, 'w');
		let offset = 0;
		while (offset < bytes.length) offset += fs.writeSync(fd, bytes, offset, bytes.length - offset);
		fs.fsyncSync(fd);
		fs.closeSync(fd);
		fd = undefined;
		fs.renameSync(tmp, file);
	} catch (error) {
		if (fd !== undefined) {
			try {
				fs.closeSync(fd);
			} catch {}
		}
		try {
			fs.unlinkSync(tmp);
		} catch {}
		throw error;
	}
}

const sameState = (a, b) => {
	const x = Y.encodeStateVector(a);
	const y = Y.encodeStateVector(b);
	return x.length === y.length && x.every((byte, i) => byte === y[i]);
};

/**
 * Re-encode a document (or an update) as a single update built on a fresh document. Returns null
 * when the result does not describe the same state: a snapshot is never replaced by a different one.
 */
export function compactUpdate(source) {
	const fresh = new Y.Doc();
	try {
		Y.applyUpdate(fresh, source instanceof Y.Doc ? Y.encodeStateAsUpdate(source) : source);
		if (source instanceof Y.Doc && !sameState(source, fresh)) return null;
		return Y.encodeStateAsUpdate(fresh);
	} finally {
		fresh.destroy();
	}
}

/**
 * Compact the snapshot file of a room that is not open. Returns `{ before, after }` in bytes, or
 * null when the file is missing, unreadable or cannot get any smaller.
 */
export function compactSnapshotFile(file, fs = realFs) {
	let before;
	try {
		before = readFileSync(file);
	} catch {
		return null;
	}
	let after;
	try {
		after = compactUpdate(before);
	} catch {
		return null;
	}
	if (!after || after.length >= before.length) return null;
	writeFileAtomic(file, after, fs);
	return { before: before.length, after: after.length };
}

/** Remove temporary files that an interrupted write left in the rooms directory. */
export function sweepRoomTemps(dir) {
	let removed = 0;
	let names = [];
	try {
		names = readdirSync(dir);
	} catch {}
	for (const name of names) {
		if (!TEMP.test(name)) continue;
		try {
			unlinkSync(join(dir, name));
			removed++;
		} catch {}
	}
	return removed;
}

/** The room names that have a snapshot file in the rooms directory. */
export function snapshotNames(dir) {
	try {
		return readdirSync(dir)
			.filter((n) => n.endsWith('.bin') && statSync(join(dir, n)).isFile())
			.map((n) => n.slice(0, -4));
	} catch {
		return [];
	}
}
