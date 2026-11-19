// Pure helpers of the reference server: loopback checks, rate limiting, awareness filtering and
// directory sizing. Kept apart from index.mjs so each stays small and testable.
import { readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';

export const LOOPBACK_HOSTS = new Set(['127.0.0.1', '::1', 'localhost', '[::1]']);
const LOOPBACK_ORIGIN = /^https?:\/\/(?:localhost|127\.0\.0\.1|\[::1\])(?::\d{1,5})?$/iu;

export const isLoopbackHost = (host) => LOOPBACK_HOSTS.has(String(host).toLowerCase());
export const isLoopbackOrigin = (origin) => LOOPBACK_ORIGIN.test(origin);

/** `decodeURIComponent` that returns null instead of throwing on bad percent-encoding. */
export function safeDecode(value) {
	try {
		return decodeURIComponent(value);
	} catch {
		return null;
	}
}

/**
 * Token bucket: `take(n)` is true while the caller stays under `rate` per second (burst `burst`).
 * `n` defaults to 1, one message; a byte budget takes the size of each message.
 */
export function createBucket(rate, burst, now = () => Date.now()) {
	let tokens = burst;
	let last = now();
	return {
		take(n = 1) {
			const t = now();
			tokens = Math.min(burst, tokens + ((t - last) / 1000) * rate);
			last = t;
			if (tokens < n) return false;
			tokens -= n;
			return true;
		},
	};
}

const utf8 = new TextEncoder();

/**
 * Keep only the awareness entries a connection may speak for: ids it already owns, or ids nobody
 * owns yet (which it then claims). `owners` maps client id to the owning connection. Returns the
 * re-encoded update (null when nothing is left) and the size in bytes of each kept state, or
 * `{ oversized: true }`, claiming nothing, when one state is longer than `maxStateBytes`.
 */
export function filterAwareness(
	update,
	conn,
	owners,
	{ maxOwned = 8, maxStateBytes = Number.POSITIVE_INFINITY } = {},
) {
	const decoder = decoding.createDecoder(update);
	const count = decoding.readVarUint(decoder);
	const entries = [];
	for (let i = 0; i < count; i++) {
		const id = decoding.readVarUint(decoder);
		const clock = decoding.readVarUint(decoder);
		const state = decoding.readVarString(decoder);
		const bytes = utf8.encode(state).length;
		if (bytes > maxStateBytes) return { oversized: true };
		entries.push({ id, clock, state, bytes });
	}
	const kept = [];
	for (const entry of entries) {
		const { id } = entry;
		const owner = owners.get(id);
		if (owner !== undefined && owner !== conn) continue;
		if (owner === undefined) {
			let owned = 0;
			for (const o of owners.values()) if (o === conn) owned++;
			if (owned >= maxOwned) continue;
			owners.set(id, conn);
		}
		kept.push(entry);
	}
	if (kept.length === 0) return { update: null, sizes: new Map() };
	const encoder = encoding.createEncoder();
	encoding.writeVarUint(encoder, kept.length);
	for (const { id, clock, state } of kept) {
		encoding.writeVarUint(encoder, id);
		encoding.writeVarUint(encoder, clock);
		encoding.writeVarString(encoder, state);
	}
	return {
		update: encoding.toUint8Array(encoder),
		sizes: new Map(kept.map(({ id, bytes }) => [id, bytes])),
	};
}

/** Total size in bytes of every file below `dir` (0 when it does not exist). */
export function directoryBytes(dir) {
	let total = 0;
	let entries;
	try {
		entries = readdirSync(dir, { withFileTypes: true });
	} catch {
		return 0;
	}
	for (const entry of entries) {
		const path = join(dir, entry.name);
		try {
			total += entry.isDirectory() ? directoryBytes(path) : statSync(path).size;
		} catch {}
	}
	return total;
}

const int = (value, fallback) => {
	const n = Number(value);
	return Number.isFinite(n) && n > 0 ? Math.floor(n) : fallback;
};
const KB = 1024;
const MB = 1024 * KB;
const GB = 1024 * MB;

/** Resolve the abuse limits from options, then the environment, then the defaults. */
export function resolveLimits(options = {}, env = process.env) {
	return {
		maxConnections: int(options.maxConnections ?? env.TEAMS_MAX_CONNECTIONS, 500),
		maxConnectionsPerIp: int(options.maxConnectionsPerIp ?? env.TEAMS_MAX_CONNECTIONS_PER_IP, 20),
		maxRooms: int(options.maxRooms ?? env.TEAMS_MAX_ROOMS, 1000),
		maxDocBytes: int(options.maxDocBytes ?? Number(env.TEAMS_MAX_DOC_MB) * MB, 64 * MB),
		messageRate: int(options.messageRate ?? env.TEAMS_MESSAGE_RATE, 100),
		messageBurst: int(options.messageBurst ?? env.TEAMS_MESSAGE_BURST, 200),
		// Per workspace (a chat room counts for the workspace it belongs to).
		workspaceQuotaBytes: int(
			options.workspaceQuotaBytes ?? Number(env.TEAMS_WORKSPACE_QUOTA_MB) * MB,
			1024 * MB,
		),
		// Per uploader, across every scope, so new scope names do not add room.
		userQuotaBytes: int(options.userQuotaBytes ?? Number(env.TEAMS_USER_QUOTA_MB) * MB, 1 * GB),
		// The whole files directory.
		filesQuotaBytes: int(options.filesQuotaBytes ?? Number(env.TEAMS_FILES_QUOTA_MB) * MB, 20 * GB),
		// Every stored room snapshot together, and how many rooms may be stored.
		roomsQuotaBytes: int(options.roomsQuotaBytes ?? Number(env.TEAMS_ROOMS_QUOTA_MB) * MB, 10 * GB),
		maxStoredRooms: int(options.maxStoredRooms ?? env.TEAMS_MAX_STORED_ROOMS, 10_000),
		// Presence: one client's state, every state of a room, and the bytes a socket sends per second.
		maxAwarenessBytes: int(
			options.maxAwarenessBytes ?? Number(env.TEAMS_MAX_AWARENESS_KB) * KB,
			64 * KB,
		),
		maxRoomAwarenessBytes: int(
			options.maxRoomAwarenessBytes ?? Number(env.TEAMS_MAX_ROOM_AWARENESS_KB) * KB,
			4 * MB,
		),
		awarenessByteRate: int(
			options.awarenessByteRate ?? Number(env.TEAMS_AWARENESS_RATE_KB) * KB,
			256 * KB,
		),
		// Refused joins, refused file requests and refused updates per client per minute.
		maxRefusals: int(options.maxRefusals ?? env.TEAMS_MAX_REFUSALS_PER_MIN, 30),
	};
}
