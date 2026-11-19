// Room snapshot storage: on-disk names, atomic writes, compaction and the sweep of leftovers. A
// snapshot is one Yjs update per room. It is written to a temporary file, flushed to disk and
// renamed over the old one, so a crash at any point leaves either the old snapshot or the new one,
// never a torn file.
import {
	closeSync,
	existsSync,
	fsyncSync,
	mkdirSync,
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
import { chatIdOfRoom, isChatRoomAlias, isCurrentChatId } from './chat-access.mjs';

const realFs = { openSync, writeSync, fsyncSync, closeSync, renameSync, unlinkSync };
const TEMP = /\.[0-9a-f]{8}\.tmp$/u;
const STORED_PREFIX = '=';
const BASE32 = 'abcdefghijklmnopqrstuvwxyz234567';
const LEGACY_NAME = /^[\w-]{1,128}$/u;
/** Where migration sets entries aside. The dot keeps it apart from every room and scope name. */
export const QUARANTINE_DIR = '.quarantine';

/**
 * The on-disk name of a room or file scope: `=` and the lower-case base32 of its exact name. Two
 * names that differ only by case stay two files on every disk, and no name can be a device name
 * such as `NUL`. The `=` tells it from the raw names of older versions, which never contain one.
 */
export function storageName(name) {
	let out = STORED_PREFIX;
	let bits = 0;
	let value = 0;
	for (const byte of Buffer.from(name, 'utf8')) {
		value = (value << 8) | byte;
		bits += 8;
		while (bits >= 5) {
			bits -= 5;
			out += BASE32[(value >>> bits) & 31];
		}
		value &= (1 << bits) - 1;
	}
	return bits > 0 ? out + BASE32[(value << (5 - bits)) & 31] : out;
}

/** The room or scope a stored name stands for, or null when it is not one {@link storageName} made. */
export function nameFromStorage(stored) {
	if (!stored.startsWith(STORED_PREFIX)) return null;
	const bytes = [];
	let bits = 0;
	let value = 0;
	for (const char of stored.slice(STORED_PREFIX.length)) {
		const digit = BASE32.indexOf(char);
		if (digit < 0) return null;
		value = (value << 5) | digit;
		bits += 5;
		if (bits >= 8) {
			bits -= 8;
			bytes.push((value >>> bits) & 255);
			value &= (1 << bits) - 1;
		}
	}
	const name = Buffer.from(bytes).toString('utf8');
	return storageName(name) === stored ? name : null;
}

const sameFile = (a, b) => {
	try {
		const [x, y] = [statSync(a), statSync(b)];
		return x.ino === y.ino && x.dev === y.dev;
	} catch {
		return false;
	}
};

/**
 * A name that is, in any spelling, the room of a chat id only the current protocol issues. No older
 * server had such a chat, so what sits under the name was put there to wait for the real one.
 */
const claimsCurrentChat = (raw) => {
	const chatId = chatIdOfRoom(raw.toLowerCase());
	return chatId !== null && isCurrentChatId(chatId);
};

/**
 * Rename the entries of `dir` that still carry a raw name (`<name><suffix>`) to their stored
 * name. Run once at start-up, before anything is served; a run that is cut short continues on the
 * next start. On a case-insensitive disk an upper-case spelling of a chat room was that chat room,
 * so it is stored under the real, lower-case name. An entry that claims a chat id of the current
 * format never becomes that room: it moves, untouched, to {@link QUARANTINE_DIR} and is logged.
 * Returns how many entries got a stored name.
 */
export function migrateStorageNames(dir, suffix, log = () => {}) {
	let entries = [];
	try {
		entries = readdirSync(dir);
	} catch {}
	let moved = 0;
	for (const entry of entries) {
		if (entry.startsWith(STORED_PREFIX) || !entry.endsWith(suffix)) continue;
		const raw = entry.slice(0, entry.length - suffix.length);
		if (!LEGACY_NAME.test(raw)) continue;
		const from = join(dir, entry);
		try {
			if (claimsCurrentChat(raw)) {
				const aside = join(dir, QUARANTINE_DIR, entry);
				mkdirSync(join(dir, QUARANTINE_DIR), { recursive: true });
				if (existsSync(aside)) throw new Error('its quarantine name is already taken');
				renameSync(from, aside);
				log(`storage: quarantined ${entry} in ${dir}: it claims a chat that did not exist yet`);
				continue;
			}
			const lower = raw.toLowerCase();
			const isSpelling = isChatRoomAlias(raw) && sameFile(join(dir, lower + suffix), from);
			const to = join(dir, storageName(isSpelling ? lower : raw) + suffix);
			if (existsSync(to)) throw new Error('its new name is already taken');
			renameSync(from, to);
			moved++;
		} catch (error) {
			log(`storage: could not migrate ${entry} (${error.message})`);
		}
	}
	if (moved) log(`storage: migrated ${moved} name(s) in ${dir}`);
	return moved;
}

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
			.map((n) => nameFromStorage(n.slice(0, -4)))
			.filter((name) => name !== null);
	} catch {
		return [];
	}
}
