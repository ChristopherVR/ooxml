// Reference bring-your-own server for the teams viewers. One process, three jobs:
//   ws  /sync/<room>      y-websocket compatible Yjs document + awareness server (persisted to disk);
//                         chat-* rooms (private chats) admit only the members listed in the room
//   ws  /signal/<room>    WebRTC signaling relay (forwards JSON to the other sockets of the room)
//   http POST|GET /files/<id>/<name>   plain file storage for attachments
//   http POST /files/link/<id>/<name>  short-lived signed download link (token in Authorization)
//   http POST /auth/session            exchange a credential for a signed session bound to one user id
//   http POST /auth/check|signout      is a session still valid / end every session of its user
//   http POST /auth/revoke             operator only (TEAMS_ADMIN_TOKEN): end every session of a user
// Nothing here understands Office formats or chat: it moves bytes, so any other server that speaks
// the same three contracts (see docs/server-contract.md) works just as well.
// Security model: see the README. One shared token opens every room; room members trust each other.
// With `authenticate` or `users`, sockets must also present a session and the user id is the one
// the session binds, never the claimed `user` parameter.
import { createServer } from 'node:http';
import { createHash, createHmac, randomBytes, timingSafeEqual } from 'node:crypto';
import {
	mkdirSync,
	existsSync,
	readFileSync,
	writeFileSync,
	createReadStream,
	createWriteStream,
	statSync,
	renameSync,
	linkSync,
	unlinkSync,
} from 'node:fs';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { WebSocketServer } from 'ws';
import * as Y from 'yjs';
import * as syncProtocol from 'y-protocols/sync';
import * as awarenessProtocol from 'y-protocols/awareness';
import * as encoding from 'lib0/encoding';
import * as decoding from 'lib0/decoding';
import {
	authenticateFromUsers,
	createSessionSigner,
	isUserId,
	plainTextUsers,
	sameSecret as sameCredential,
} from './auth.mjs';
import { createSlidingLimiter } from './rate-limit.mjs';
import { authorizeChatJoin, isChatRoom, roomMembers } from './chat-access.mjs';
import {
	createBucket,
	directoryBytes,
	filterAwareness,
	isLoopbackHost,
	isLoopbackOrigin,
	resolveLimits,
	safeDecode,
} from './guards.mjs';

const MESSAGE_SYNC = 0;
const MESSAGE_AWARENESS = 1;
const MESSAGE_QUERY_AWARENESS = 3;
const ROOM = /^[\w-]{1,128}$/u;

/**
 * What `authorizeSync` returns: `'write'`, `'read'`, `{ access, principal? }`, or a falsy value
 * to refuse. A principal (default: the session or claimed user) names the socket for `closeSync`.
 * @typedef {'write' | 'read' | { access: 'write' | 'read', principal?: string } | null | undefined | false} SyncVerdict
 */

/** Normalise an `authorizeSync` verdict; null refuses the socket. */
export function syncGrant(verdict, user) {
	const access = typeof verdict === 'string' ? verdict : verdict?.access;
	if (access !== 'write' && access !== 'read') return null;
	const principal =
		typeof verdict === 'object' && typeof verdict.principal === 'string' ? verdict.principal : user;
	return { access, principal: principal ?? null };
}
const WS_OPEN = 1;
const SERVER_OWNER = Symbol('server');

const sameSecret = sameCredential;

const unlinkQuiet = (path) => {
	try {
		unlinkSync(path);
	} catch {}
};

/**
 * @param {{ port?: number, host?: string, token?: string, origins?: string[], dataDir?: string,
 *   insecure?: boolean, authenticate?: (request: { user: string, credential?: string,
 *   headers: Record<string, string | string[] | undefined> }) => unknown, users?: Record<string, string>,
 *   sessionSecret?: string, sessionTtlMs?: number, adminToken?: string,
 *   sessionRate?: { windowMs?: number, perIp?: number, perUser?: number }, maxFileBytes?: number, maxSyncBytes?: number, maxSignalBytes?: number,
 *   maxConnections?: number, maxConnectionsPerIp?: number, maxRooms?: number, maxDocBytes?: number,
 *   messageRate?: number, messageBurst?: number, filesQuotaBytes?: number,
 *   authorizeSync?: (request: { room: string, user: string | null, url: URL,
 *   headers: Record<string, string | string[] | undefined> }) => SyncVerdict | Promise<SyncVerdict>,
 *   log?: (m: string) => void }} [options]
 */
export function createTeamsServer(options = {}) {
	const {
		token = process.env.TEAMS_TOKEN || '',
		origins = (process.env.TEAMS_ORIGINS || '')
			.split(',')
			.map((s) => s.trim())
			.filter(Boolean),
		dataDir = resolve(process.env.TEAMS_DATA || './data'),
		insecure = process.env.TEAMS_INSECURE === '1',
		users,
		sessionSecret = process.env.TEAMS_SESSION_SECRET || '',
		sessionTtlMs,
		adminToken = process.env.TEAMS_ADMIN_TOKEN || '',
		maxFileBytes = 25 * 1024 * 1024,
		maxSyncBytes = 32 * 1024 * 1024,
		maxSignalBytes = 256 * 1024,
		log = () => {},
	} = options;
	const authenticate = options.authenticate ?? (users ? authenticateFromUsers(users) : undefined);
	// Per-room permissions decided by the host (see "Per-room permissions" in the README).
	const authorizeSync = typeof options.authorizeSync === 'function' ? options.authorizeSync : null;
	// Real accounts exist: sockets must carry a session. Otherwise identity is a claim (dev default).
	const authRequired = typeof authenticate === 'function';
	if (users) {
		const plain = plainTextUsers(users).length;
		if (plain > 0)
			log(
				`WARNING: ${plain} user credential(s) are stored as plain text; hash them with openteams-hash-password`,
			);
	}
	// Session revocation: a per-user generation that tokens carry, persisted next to the data so a
	// revocation survives a restart that keeps the same session secret.
	const revocationsFile = join(dataDir, 'session-revocations.json');
	const loadGenerations = () => {
		try {
			const parsed = JSON.parse(readFileSync(revocationsFile, 'utf8'));
			return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
		} catch {
			return {};
		}
	};
	const signer = createSessionSigner({
		secret: sessionSecret,
		...(sessionTtlMs ? { ttlMs: sessionTtlMs } : {}),
		generations: loadGenerations(),
		onRevoke: (user, _generation, all) => {
			try {
				writeFileSync(revocationsFile, JSON.stringify(all));
			} catch (error) {
				log(`could not persist session revocations (${error.message})`);
			}
		},
	});
	const envInt = (name) => Number(process.env[name]) || undefined;
	const rate = {
		windowMs: options.sessionRate?.windowMs ?? envInt('TEAMS_AUTH_RATE_WINDOW_MS') ?? 60_000,
		perIp: options.sessionRate?.perIp ?? envInt('TEAMS_AUTH_RATE_PER_IP') ?? 30,
		perUser: options.sessionRate?.perUser ?? envInt('TEAMS_AUTH_RATE_PER_USER') ?? 10,
	};
	const ipLimiter = createSlidingLimiter({ limit: rate.perIp, windowMs: rate.windowMs });
	const userLimiter = createSlidingLimiter({ limit: rate.perUser, windowMs: rate.windowMs });
	/** Open sockets that were admitted with a session, so a revocation can drop them. */
	const sessionSockets = new Map();
	const revokeSessions = (user) => {
		signer.revoke(user);
		for (const [ws, owner] of sessionSockets) if (owner === user) ws.close(1008, 'session revoked');
	};
	const limits = resolveLimits(options);
	const roomsDir = join(dataDir, 'rooms');
	const filesDir = join(dataDir, 'files');
	mkdirSync(roomsDir, { recursive: true });
	mkdirSync(filesDir, { recursive: true });
	let filesBytes = directoryBytes(filesDir);

	// Explicit origins win. Without them and without a token, only loopback pages and non-browser
	// clients (no Origin header) get in, so a random website cannot drive a local server.
	const originAllowed = (origin) => {
		if (origins.length > 0) return Boolean(origin) && origins.includes(origin);
		if (token) return true;
		return !origin || isLoopbackOrigin(origin);
	};
	const tokenOk = (url) => !token || sameSecret(url.searchParams.get('token') ?? '', token);

	// ---- signed file links ------------------------------------------------------------------
	// A download link is valid for one path and a few minutes, so the long-lived token never has to
	// appear in a URL (history, logs, Referer). WebSockets cannot send headers, so only /sync and
	// /signal still take `?token=`.
	const LINK_TTL_MS = 5 * 60 * 1000;
	const linkKey = createHash('sha256').update(`teams-file-link|${token}`).digest();
	const sign = (path, exp) => createHmac('sha256', linkKey).update(`${path}|${exp}`).digest('hex');
	const signedUrl = (path) => {
		const exp = Date.now() + LINK_TTL_MS;
		return `${path}?sig=${sign(path, exp)}&exp=${exp}`;
	};
	const linkOk = (url) => {
		const exp = Number(url.searchParams.get('exp'));
		const sig = url.searchParams.get('sig') ?? '';
		return Number.isFinite(exp) && exp > Date.now() && sameSecret(sig, sign(url.pathname, exp));
	};

	// ---- sync rooms -------------------------------------------------------------------------
	/** @type {Map<string, any>} */
	const syncRooms = new Map();
	const snapshotPath = (name) => join(roomsDir, `${name}.bin`);

	const send = (ws, message) => {
		if (ws.readyState === WS_OPEN) ws.send(message, (err) => err && ws.close());
	};

	/** Write a room snapshot; a full disk or a permission error is logged, never fatal. */
	function flushRoom(name, room) {
		const file = snapshotPath(name);
		const tmp = `${file}.${randomBytes(4).toString('hex')}.tmp`;
		try {
			writeFileSync(tmp, Y.encodeStateAsUpdate(room.doc));
			renameSync(tmp, file);
		} catch (error) {
			unlinkQuiet(tmp);
			log(`room ${name}: could not persist snapshot (${error.message})`);
		}
	}

	/** The members stored for a chat room, read without keeping a room that nobody joined. */
	function membersOf(name) {
		const live = syncRooms.get(name);
		if (live) return roomMembers(live.doc);
		const file = snapshotPath(name);
		if (!existsSync(file)) return new Set();
		const doc = new Y.Doc();
		try {
			Y.applyUpdate(doc, readFileSync(file));
			return roomMembers(doc);
		} catch {
			return new Set();
		} finally {
			doc.destroy();
		}
	}

	/** Close the sockets of people who are no longer listed as members of a chat room. */
	function enforceMembers(room) {
		const members = roomMembers(room.doc);
		if (members.size === 0) return;
		for (const [ws, user] of room.users) if (!members.has(user)) ws.close(1008, 'not a member');
	}

	function getRoom(name) {
		let room = syncRooms.get(name);
		if (room) {
			clearTimeout(room.idle);
			return room;
		}
		const doc = new Y.Doc();
		const file = snapshotPath(name);
		let bytes = 0;
		if (existsSync(file)) {
			try {
				const snapshot = readFileSync(file);
				Y.applyUpdate(doc, snapshot);
				bytes = snapshot.length;
			} catch (error) {
				log(`room ${name}: ignoring unreadable snapshot (${error.message})`);
			}
		}
		const awareness = new awarenessProtocol.Awareness(doc);
		awareness.setLocalState(null);
		if (isChatRoom(name)) doc.on('update', () => enforceMembers(room));
		room = {
			doc,
			awareness,
			conns: new Map(),
			owners: new Map([[awareness.clientID, SERVER_OWNER]]),
			/** socket -> the user it joined as (chat rooms only) */
			users: new Map(),
			/** socket -> { access, principal } from `authorizeSync` (absent: full access) */
			grants: new Map(),
			bytes,
			timer: null,
			idle: null,
		};
		const persist = () => {
			clearTimeout(room.timer);
			room.timer = setTimeout(() => flushRoom(name, room), 500);
			room.timer.unref?.();
		};
		doc.on('update', (update, origin) => {
			const encoder = encoding.createEncoder();
			encoding.writeVarUint(encoder, MESSAGE_SYNC);
			syncProtocol.writeUpdate(encoder, update);
			const message = encoding.toUint8Array(encoder);
			for (const conn of room.conns.keys()) if (conn !== origin) send(conn, message);
			persist();
		});
		awareness.on('update', ({ added, updated, removed }, origin) => {
			const changed = [...added, ...updated, ...removed];
			if (origin && room.conns.has(origin)) {
				const owned = room.conns.get(origin);
				for (const id of added) owned.add(id);
				for (const id of removed) {
					owned.delete(id);
					if (room.owners.get(id) === origin) room.owners.delete(id);
				}
			}
			const encoder = encoding.createEncoder();
			encoding.writeVarUint(encoder, MESSAGE_AWARENESS);
			encoding.writeVarUint8Array(
				encoder,
				awarenessProtocol.encodeAwarenessUpdate(awareness, changed),
			);
			const message = encoding.toUint8Array(encoder);
			for (const conn of room.conns.keys()) send(conn, message);
		});
		syncRooms.set(name, room);
		return room;
	}

	function leaveSync(name, room, ws) {
		const owned = room.conns.get(ws);
		room.users.delete(ws);
		room.grants.delete(ws);
		room.conns.delete(ws);
		for (const [id, owner] of room.owners) if (owner === ws) room.owners.delete(id);
		if (owned?.size) awarenessProtocol.removeAwarenessStates(room.awareness, [...owned], null);
		if (room.conns.size === 0) {
			room.idle = setTimeout(() => {
				if (room.conns.size > 0) return;
				clearTimeout(room.timer);
				flushRoom(name, room);
				room.awareness.destroy();
				room.doc.destroy();
				syncRooms.delete(name);
			}, 30_000);
			room.idle.unref?.();
		}
	}

	function handleSync(ws, name, admission, grant) {
		const room = getRoom(name);
		if (grant) room.grants.set(ws, grant);
		const readOnly = grant?.access === 'read';
		if (admission) {
			// A new chat room starts with the members its first socket listed.
			if (admission.seed.length > 0 && roomMembers(room.doc).size === 0)
				room.doc.transact(() => {
					for (const id of admission.seed) room.doc.getMap('teams:chat-members').set(id, true);
				});
			room.users.set(ws, admission.user);
		}
		room.conns.set(ws, new Set());
		ws.binaryType = 'arraybuffer';
		const bucket = createBucket(limits.messageRate, limits.messageBurst);
		ws.on('message', (data) => {
			if (!bucket.take()) return void ws.close(1008, 'rate limit');
			try {
				const decoder = decoding.createDecoder(new Uint8Array(data));
				const type = decoding.readVarUint(decoder);
				if (type === MESSAGE_SYNC) {
					// A read-only socket may ask for the document (step 1), never change it.
					if (readOnly && decoding.peekVarUint(decoder) !== syncProtocol.messageYjsSyncStep1)
						return;
					if (room.bytes + data.byteLength > limits.maxDocBytes) {
						// The running total over-counts (merged and deleted content shrinks); measure once.
						room.bytes = Y.encodeStateAsUpdate(room.doc).length;
						if (room.bytes + data.byteLength > limits.maxDocBytes) {
							log(`sync ${name}: document is at its size limit, update rejected`);
							return void ws.close(1009, 'document too large');
						}
					}
					room.bytes += data.byteLength;
					const encoder = encoding.createEncoder();
					encoding.writeVarUint(encoder, MESSAGE_SYNC);
					syncProtocol.readSyncMessage(decoder, encoder, room.doc, ws);
					if (encoding.length(encoder) > 1) send(ws, encoding.toUint8Array(encoder));
				} else if (type === MESSAGE_AWARENESS) {
					// A connection may only speak for the client ids it announced itself.
					const update = filterAwareness(decoding.readVarUint8Array(decoder), ws, room.owners);
					if (update) awarenessProtocol.applyAwarenessUpdate(room.awareness, update, ws);
				} else if (type === MESSAGE_QUERY_AWARENESS) {
					const states = [...room.awareness.getStates().keys()];
					if (states.length) {
						const encoder = encoding.createEncoder();
						encoding.writeVarUint(encoder, MESSAGE_AWARENESS);
						encoding.writeVarUint8Array(
							encoder,
							awarenessProtocol.encodeAwarenessUpdate(room.awareness, states),
						);
						send(ws, encoding.toUint8Array(encoder));
					}
				}
			} catch (error) {
				log(`sync ${name}: dropped malformed message (${error.message})`);
			}
		});
		ws.on('close', () => leaveSync(name, room, ws));
		ws.on('error', () => ws.close());
		// Our side of the handshake: ask for the client's state and offer ours.
		const hello = encoding.createEncoder();
		encoding.writeVarUint(hello, MESSAGE_SYNC);
		syncProtocol.writeSyncStep1(hello, room.doc);
		send(ws, encoding.toUint8Array(hello));
		const states = [...room.awareness.getStates().keys()];
		if (states.length) {
			const encoder = encoding.createEncoder();
			encoding.writeVarUint(encoder, MESSAGE_AWARENESS);
			encoding.writeVarUint8Array(
				encoder,
				awarenessProtocol.encodeAwarenessUpdate(room.awareness, states),
			);
			send(ws, encoding.toUint8Array(encoder));
		}
	}

	// ---- signaling rooms --------------------------------------------------------------------
	/** @type {Map<string, Map<any, string | null>>} socket -> the peer id it speaks for */
	const signalRooms = new Map();
	function handleSignal(ws, name) {
		let peers = signalRooms.get(name);
		if (!peers) signalRooms.set(name, (peers = new Map()));
		peers.set(ws, null);
		let peerId = null;
		const bucket = createBucket(limits.messageRate, limits.messageBurst);
		ws.on('message', (data, isBinary) => {
			if (!bucket.take()) return void ws.close(1008, 'rate limit');
			if (isBinary) return;
			const text = data.toString();
			let parsed;
			try {
				parsed = JSON.parse(text);
			} catch {
				return;
			}
			if (!parsed || typeof parsed.from !== 'string') return;
			if (parsed.to !== undefined && typeof parsed.to !== 'string') return;
			if (peerId === null) {
				// A socket speaks for one peer id: the first one it claims, if nobody else holds it.
				for (const [other, id] of peers)
					if (other !== ws && id === parsed.from) return void ws.close(1008, 'peer id in use');
				peerId = parsed.from;
				peers.set(ws, peerId);
			}
			if (parsed.from !== peerId) return;
			for (const [other, id] of peers) {
				if (other === ws) continue;
				if (parsed.to === undefined || id === parsed.to) send(other, text);
			}
		});
		ws.on('close', () => {
			peers.delete(ws);
			if (peerId)
				for (const other of peers.keys())
					send(other, JSON.stringify({ type: 'bye', from: peerId }));
			if (peers.size === 0) signalRooms.delete(name);
		});
		ws.on('error', () => ws.close());
	}

	// ---- files ------------------------------------------------------------------------------
	const cors = (res, req) => {
		const origin = req.headers.origin;
		if (origin && originAllowed(origin)) {
			res.setHeader('Access-Control-Allow-Origin', origin);
			res.setHeader('Vary', 'Origin');
		}
		res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
		res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
		res.setHeader('Access-Control-Expose-Headers', 'Retry-After');
	};
	/** The stored file name, or null for bad percent-encoding and names made only of dots. */
	const cleanName = (n) => {
		const decoded = safeDecode(n);
		if (decoded === null) return null;
		const name = decoded.replace(/[^\w.\- ()]/gu, '_').slice(0, 120);
		return name === '' || /^\.+$/u.test(name) ? null : name;
	};
	const plain = (res, status, headers = {}) => void res.writeHead(status, headers).end();

	function download(res, file) {
		let stat;
		try {
			stat = statSync(file);
		} catch {
			return plain(res, 404);
		}
		if (!stat.isFile()) return plain(res, 404);
		const stream = createReadStream(file);
		stream.on('error', () => (res.headersSent ? res.destroy() : plain(res, 500)));
		stream.on('open', () => {
			res.writeHead(200, {
				'Content-Type': 'application/octet-stream',
				'Content-Length': stat.size,
				'Content-Disposition': 'attachment',
				'Referrer-Policy': 'no-referrer',
				'X-Content-Type-Options': 'nosniff',
			});
			stream.pipe(res);
		});
		res.on('close', () => stream.destroy());
	}

	function upload(req, res, dir, name, id) {
		const declared = Number(req.headers['content-length']);
		const refuse = (status) => {
			plain(res, status, { Connection: 'close' });
			res.once('finish', () => req.destroy());
		};
		if (Number.isFinite(declared) && declared > maxFileBytes) return refuse(413);
		if (filesBytes + (Number.isFinite(declared) ? declared : 0) > limits.filesQuotaBytes)
			return refuse(507);
		mkdirSync(dir, { recursive: true });
		const tmp = join(dir, `.${randomBytes(6).toString('hex')}.part`);
		const out = createWriteStream(tmp, { flags: 'wx' });
		let size = 0;
		let failed = false;
		const abort = (status) => {
			if (failed) return;
			failed = true;
			req.unpipe(out);
			out.destroy();
			if (!res.headersSent) refuse(status);
		};
		out.on('close', () => failed && unlinkQuiet(tmp));
		req.on('data', (chunk) => {
			size += chunk.length;
			if (size > maxFileBytes) abort(413);
			else if (filesBytes + size > limits.filesQuotaBytes) abort(507);
		});
		req.on('error', () => abort(400));
		req.on('aborted', () => abort(400));
		out.on('error', (error) => {
			log(`file ${id}: write failed (${error.message})`);
			abort(500);
		});
		req.pipe(out);
		out.on('finish', () => {
			if (failed) return;
			// Never overwrite: link() fails when the name exists, so pick a fresh one.
			const dot = name.lastIndexOf('.');
			const [stem, ext] = dot > 0 ? [name.slice(0, dot), name.slice(dot)] : [name, ''];
			let finalName = name;
			for (let attempt = 0; ; attempt++) {
				try {
					linkSync(tmp, join(dir, finalName));
					break;
				} catch (error) {
					if (error.code !== 'EEXIST' || attempt >= 5) {
						failed = true;
						unlinkQuiet(tmp);
						log(`file ${id}: could not store ${name} (${error.message})`);
						return plain(res, 500);
					}
					finalName = `${stem.slice(0, 100)}-${randomBytes(3).toString('hex')}${ext}`;
				}
			}
			unlinkQuiet(tmp);
			filesBytes += size;
			res
				.writeHead(201, { 'Content-Type': 'application/json' })
				.end(JSON.stringify({ url: `/files/${id}/${encodeURIComponent(finalName)}`, size }));
		});
	}

	// ---- sessions ---------------------------------------------------------------------------
	const json = (res, status, body) =>
		void res
			.writeHead(status, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' })
			.end(JSON.stringify(body));

	/** POST /auth/session {user, credential}: the workspace token (if any) in Authorization. */
	function sessionRequest(req, res) {
		if (req.method !== 'POST') return plain(res, 405);
		if (!originAllowed(req.headers.origin ?? '')) return plain(res, 403);
		const bearer = /^Bearer (.+)$/u.exec(req.headers.authorization ?? '')?.[1];
		if (token && (bearer === undefined || !sameSecret(bearer, token))) return plain(res, 401);
		const ipVerdict = ipLimiter.check(req.socket.remoteAddress ?? '');
		if (!ipVerdict.ok) return tooMany(res, ipVerdict);
		readSmallJson(req, res, async (body) => {
			const user = body?.user;
			const credential = typeof body?.credential === 'string' ? body.credential : undefined;
			if (!isUserId(user)) return plain(res, 400);
			const userVerdict = userLimiter.check(user);
			if (!userVerdict.ok) return tooMany(res, userVerdict);
			try {
				// Dev default (no hook): the claimed id is accepted. Not secure, see the README.
				const verdict = authRequired
					? await authenticate({
							user,
							...(credential === undefined ? {} : { credential }),
							headers: req.headers,
						})
					: { id: user };
				const id = typeof verdict === 'string' ? verdict : verdict?.id;
				if (!isUserId(id)) return plain(res, 401);
				const issued = signer.issue(id);
				json(res, 200, { session: issued.token, user: id, expiresAt: issued.expiresAt });
			} catch (error) {
				// Never log the credential or the hook's input, only the failure.
				log(`authenticate failed (${error?.message ?? 'error'})`);
				plain(res, 500);
			}
		});
	}

	const tooMany = (res, verdict) =>
		plain(res, 429, { 'Retry-After': String(verdict.retryAfterSec), 'Cache-Control': 'no-store' });

	/** Read a small JSON body (at most 4 KiB), then call `done(body)`; 400 when it is not JSON. */
	function readSmallJson(req, res, done) {
		const chunks = [];
		let size = 0;
		req.on('data', (chunk) => {
			size += chunk.length;
			if (size > 4096) {
				req.destroy();
				return;
			}
			chunks.push(chunk);
		});
		req.on('error', () => {});
		req.on('end', async () => {
			let body;
			try {
				body = JSON.parse(Buffer.concat(chunks).toString('utf8'));
			} catch {
				return plain(res, 400);
			}
			done(body);
		});
	}

	/** POST /auth/signout {session}: revoke every session of the user the token belongs to. */
	function signoutRequest(req, res) {
		if (req.method !== 'POST') return plain(res, 405);
		if (!originAllowed(req.headers.origin ?? '')) return plain(res, 403);
		const bearer = /^Bearer (.+)$/u.exec(req.headers.authorization ?? '')?.[1];
		if (token && (bearer === undefined || !sameSecret(bearer, token))) return plain(res, 401);
		readSmallJson(req, res, (body) => {
			const user = signer.verify(body?.session);
			if (user === null) return plain(res, 401);
			revokeSessions(user);
			plain(res, 204);
		});
	}

	/** POST /auth/check {session}: 204 while the session is valid, 401 when it is not. */
	function checkRequest(req, res) {
		if (req.method !== 'POST') return plain(res, 405);
		if (!originAllowed(req.headers.origin ?? '')) return plain(res, 403);
		const bearer = /^Bearer (.+)$/u.exec(req.headers.authorization ?? '')?.[1];
		if (token && (bearer === undefined || !sameSecret(bearer, token))) return plain(res, 401);
		readSmallJson(req, res, (body) =>
			plain(res, signer.verify(body?.session) === null ? 401 : 204),
		);
	}

	/** POST /auth/revoke {user}: admin only (TEAMS_ADMIN_TOKEN as the Bearer token). */
	function revokeRequest(req, res) {
		if (!adminToken) return plain(res, 404);
		if (req.method !== 'POST') return plain(res, 405);
		const bearer = /^Bearer (.+)$/u.exec(req.headers.authorization ?? '')?.[1];
		if (bearer === undefined || !sameSecret(bearer, adminToken)) return plain(res, 401);
		const ipVerdict = ipLimiter.check(req.socket.remoteAddress ?? '');
		if (!ipVerdict.ok) return tooMany(res, ipVerdict);
		readSmallJson(req, res, (body) => {
			if (!isUserId(body?.user)) return plain(res, 400);
			revokeSessions(body.user);
			plain(res, 204);
		});
	}

	function onRequest(req, res) {
		cors(res, req);
		const url = new URL(req.url ?? '/', 'http://localhost');
		if (req.method === 'OPTIONS') return plain(res, 204);
		if (url.pathname === '/health') {
			return void res.writeHead(200, { 'Content-Type': 'application/json' }).end(
				JSON.stringify({
					ok: true,
					rooms: syncRooms.size,
					calls: signalRooms.size,
					auth: Boolean(token),
					accounts: authRequired,
				}),
			);
		}
		if (url.pathname === '/auth/session') return void sessionRequest(req, res);
		if (url.pathname === '/auth/check') return void checkRequest(req, res);
		if (url.pathname === '/auth/signout') return void signoutRequest(req, res);
		if (url.pathname === '/auth/revoke') return void revokeRequest(req, res);
		const link = /^\/files\/link\/([\w-]{1,64})\/([^/]{1,200})$/u.exec(url.pathname);
		const m = link ?? /^\/files\/([\w-]{1,64})\/([^/]{1,200})$/u.exec(url.pathname);
		if (!m) return plain(res, 404);
		if (!originAllowed(req.headers.origin ?? '')) return plain(res, 403);
		const bearer = /^Bearer (.+)$/u.exec(req.headers.authorization ?? '')?.[1];
		const authorised = bearer !== undefined && sameSecret(bearer, token);
		// A signed link only ever opens its own file, and only for reading.
		const signedRead = req.method === 'GET' && !link && linkOk(url);
		if (token && !authorised && !signedRead) return plain(res, 401);
		const name = cleanName(m[2]);
		if (name === null) return plain(res, 400);
		if (link) {
			if (req.method !== 'POST') return plain(res, 405);
			const path = `/files/${m[1]}/${m[2]}`;
			return void res
				.writeHead(200, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' })
				.end(JSON.stringify({ url: token ? signedUrl(path) : path }));
		}
		const dir = join(filesDir, m[1]);
		if (req.method === 'GET') return download(res, join(dir, name));
		if (req.method === 'POST') return upload(req, res, dir, name, m[1]);
		plain(res, 405);
	}

	const http = createServer((req, res) => {
		try {
			onRequest(req, res);
		} catch (error) {
			log(`request failed (${error.message})`);
			if (res.headersSent) res.destroy();
			else plain(res, 500);
		}
	});

	// ---- upgrades ---------------------------------------------------------------------------
	const syncWss = new WebSocketServer({ noServer: true, maxPayload: maxSyncBytes });
	const signalWss = new WebSocketServer({ noServer: true, maxPayload: maxSignalBytes });
	/** @type {Map<string, number>} open sockets per remote address */
	const perIp = new Map();
	const reject = (socket, line) => {
		socket.write(`HTTP/1.1 ${line}\r\nConnection: close\r\n\r\n`);
		socket.destroy();
	};
	function onUpgrade(req, socket, head) {
		const url = new URL(req.url ?? '/', 'http://localhost');
		const m = /^\/(sync|signal)\/([^/]+)$/u.exec(url.pathname);
		const name = m ? safeDecode(m[2]) : null;
		if (!m) return void socket.destroy();
		if (name === null || !ROOM.test(name)) return reject(socket, '400 Bad Request');
		if (!originAllowed(req.headers.origin ?? '')) return reject(socket, '403 Forbidden');
		if (!tokenOk(url)) return reject(socket, '401 Unauthorized');
		// Identity: with accounts the session decides who this is; without, the claim stands (dev).
		const sessionUser = signer.verify(url.searchParams.get('session'));
		if (url.searchParams.has('session') && sessionUser === null)
			return reject(socket, '401 Unauthorized');
		if (authRequired && sessionUser === null) return reject(socket, '401 Unauthorized');
		// Private chats: only the people listed in the room's member map may join.
		let admission = null;
		if (m[1] === 'sync' && isChatRoom(name)) {
			const user = sessionUser ?? url.searchParams.get('user') ?? '';
			const verdict = authorizeChatJoin({
				name,
				user,
				members: url.searchParams.get('members'),
				existing: membersOf(name),
			});
			if (!verdict.ok) return reject(socket, '403 Forbidden');
			admission = { seed: verdict.seed, user };
		}
		const ip = req.socket.remoteAddress ?? '';
		const open = syncWss.clients.size + signalWss.clients.size;
		const rooms = m[1] === 'sync' ? syncRooms : signalRooms;
		if (
			open >= limits.maxConnections ||
			(perIp.get(ip) ?? 0) >= limits.maxConnectionsPerIp ||
			(!rooms.has(name) && rooms.size >= limits.maxRooms)
		)
			return reject(socket, '503 Service Unavailable');
		const admit = (grant) => {
			if (socket.destroyed) return;
			perIp.set(ip, (perIp.get(ip) ?? 0) + 1);
			upgradeSocket(req, socket, head, m[1], name, admission, sessionUser, ip, grant);
		};
		if (m[1] !== 'sync' || !authorizeSync) return admit(null);
		const user = sessionUser ?? url.searchParams.get('user');
		Promise.resolve()
			.then(() => authorizeSync({ room: name, user, url, headers: req.headers }))
			.then(
				(verdict) => {
					const grant = syncGrant(verdict, user);
					if (!grant) return reject(socket, '403 Forbidden');
					return admit(grant);
				},
				(error) => {
					log(`sync ${name}: authorizeSync failed (${error?.message ?? error})`);
					reject(socket, '500 Internal Server Error');
				},
			);
	}
	function upgradeSocket(req, socket, head, kind, name, admission, sessionUser, ip, grant) {
		const wss = kind === 'sync' ? syncWss : signalWss;
		wss.handleUpgrade(req, socket, head, (ws) => {
			if (sessionUser !== null) {
				sessionSockets.set(ws, sessionUser);
				ws.on('close', () => sessionSockets.delete(ws));
			}
			ws.on('close', () => {
				const left = (perIp.get(ip) ?? 1) - 1;
				if (left <= 0) perIp.delete(ip);
				else perIp.set(ip, left);
			});
			try {
				if (kind === 'sync') handleSync(ws, name, admission, grant);
				else handleSignal(ws, name);
			} catch (error) {
				log(`${kind} ${name}: setup failed (${error.message})`);
				ws.terminate();
			}
		});
	}
	const upgrade = (req, socket, head) => {
		socket.on('error', () => socket.destroy());
		try {
			onUpgrade(req, socket, head);
		} catch (error) {
			log(`upgrade failed (${error.message})`);
			socket.destroy();
		}
	};
	http.on('upgrade', upgrade);

	return {
		http,
		/**
		 * Handle a `/sync/<room>` or `/signal/<room>` upgrade that arrived on another HTTP server,
		 * so a host can serve these sockets from its own origin. Other paths are refused.
		 */
		handleUpgrade: upgrade,
		/**
		 * Close the sync sockets of `room` (only those whose `authorizeSync` principal is
		 * `principal`, when given) so they reconnect and are authorised again. Returns the count.
		 */
		closeSync(room, principal, code = 4403, reason = 'access changed') {
			const live = syncRooms.get(room);
			let closed = 0;
			for (const ws of live?.conns.keys() ?? []) {
				if (principal !== undefined && live.grants.get(ws)?.principal !== principal) continue;
				ws.close(code, reason);
				closed++;
			}
			return closed;
		},
		/** The distinct `authorizeSync` principals of the sockets open in a sync room. */
		syncPrincipals(room) {
			const grants = [...(syncRooms.get(room)?.grants.values() ?? [])];
			return [...new Set(grants.map((grant) => grant.principal).filter(Boolean))];
		},
		/** Whether a sync room holds no document content yet, live or persisted. */
		isSyncRoomEmpty(room) {
			const live = syncRooms.get(room);
			if (live) return live.doc.store.clients.size === 0;
			if (!ROOM.test(room)) return true;
			const file = snapshotPath(room);
			if (!existsSync(file)) return true;
			const doc = new Y.Doc();
			try {
				Y.applyUpdate(doc, readFileSync(file));
				return doc.store.clients.size === 0;
			} catch {
				return true;
			} finally {
				doc.destroy();
			}
		},
		accounts: authRequired,
		/** Invalidate every session issued for `user` and drop their open sockets. */
		revokeSessions,
		listen: (port = Number(process.env.PORT) || 8787, host = process.env.HOST || '127.0.0.1') => {
			if (!token && !isLoopbackHost(host)) {
				if (!insecure)
					return Promise.reject(
						new Error(
							`refusing to listen on ${host} without a token: set TEAMS_TOKEN, or TEAMS_INSECURE=1 (--insecure) to accept that anyone who can reach it can read and write everything`,
						),
					);
				log(`WARNING: listening on ${host} with no token and TEAMS_INSECURE set`);
			}
			return new Promise((ok, fail) => {
				http.once('error', fail);
				http.listen(port, host, () => {
					http.off('error', fail);
					ok(http.address());
				});
			});
		},
		close: () =>
			new Promise((ok) => {
				// Flush every room (a debounced snapshot must not be lost on shutdown), then release it.
				for (const [name, room] of syncRooms) {
					clearTimeout(room.timer);
					clearTimeout(room.idle);
					flushRoom(name, room);
					room.awareness.destroy();
					room.doc.destroy();
				}
				syncRooms.clear();
				for (const c of [...syncWss.clients, ...signalWss.clients]) c.terminate();
				http.close(() => ok());
				http.closeAllConnections();
			}),
	};
}

/** Start a server configured from the environment and log where it listens (the CLI entry). */
export async function runTeamsServer() {
	const insecure = process.env.TEAMS_INSECURE === '1' || process.argv.includes('--insecure');
	// TEAMS_USERS_FILE: a JSON object of user id -> credential (see the README).
	const usersFile = process.env.TEAMS_USERS_FILE;
	let users;
	if (usersFile) {
		try {
			users = JSON.parse(readFileSync(usersFile, 'utf8'));
		} catch {
			// The parser's message can quote the file, which holds credentials: say nothing of it.
			throw new Error('TEAMS_USERS_FILE could not be read as a JSON object');
		}
		if (!users || typeof users !== 'object' || Array.isArray(users))
			throw new Error('TEAMS_USERS_FILE could not be read as a JSON object');
	}
	const server = createTeamsServer({
		insecure,
		...(users ? { users } : {}),
		log: (m) => console.log(m),
	});
	const addr = await server.listen();
	const auth = process.env.TEAMS_TOKEN
		? 'token required'
		: 'no token (loopback only unless TEAMS_INSECURE=1)';
	const accounts = server.accounts
		? 'accounts required'
		: 'NO ACCOUNTS: user ids are claims (dev default, not secure)';
	console.log(`teams server on http://${addr.address}:${addr.port}  [${auth}; ${accounts}]`);
	console.log(
		'  ws://%s:%d/sync/<room>   ws://%s:%d/signal/<room>   /files/<id>/<name>',
		addr.address,
		addr.port,
		addr.address,
		addr.port,
	);
	return server;
}

/** CLI wrapper: print a startup refusal plainly and exit non-zero instead of a stack trace. */
export async function runTeamsServerCli() {
	try {
		await runTeamsServer();
	} catch (error) {
		console.error(`openteams-server: ${error.message}`);
		process.exit(1);
	}
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
	await runTeamsServerCli();
}
