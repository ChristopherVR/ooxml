// Reference bring-your-own server for the teams viewers. One process, three jobs:
//   ws  /sync/<room>      y-websocket compatible Yjs document + awareness server (persisted to disk)
//   ws  /signal/<room>    WebRTC signaling relay (forwards JSON to the other sockets of the room)
//   http POST|GET /files/<id>/<name>   plain file storage for attachments
//   http POST /files/link/<id>/<name>  short-lived signed download link (token in Authorization)
// Nothing here understands Office formats or chat: it moves bytes, so any other server that speaks
// the same three contracts (see docs/server-contract.md) works just as well.
// Security model: see the README. One shared token opens every room; room members trust each other.
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
const WS_OPEN = 1;
const SERVER_OWNER = Symbol('server');

const sameSecret = (a, b) => {
	const ha = createHash('sha256').update(String(a)).digest();
	const hb = createHash('sha256').update(String(b)).digest();
	return timingSafeEqual(ha, hb);
};

const unlinkQuiet = (path) => {
	try {
		unlinkSync(path);
	} catch {}
};

/**
 * @param {{ port?: number, host?: string, token?: string, origins?: string[], dataDir?: string,
 *   insecure?: boolean, maxFileBytes?: number, maxSyncBytes?: number, maxSignalBytes?: number,
 *   maxConnections?: number, maxConnectionsPerIp?: number, maxRooms?: number, maxDocBytes?: number,
 *   messageRate?: number, messageBurst?: number, filesQuotaBytes?: number,
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
		maxFileBytes = 25 * 1024 * 1024,
		maxSyncBytes = 32 * 1024 * 1024,
		maxSignalBytes = 256 * 1024,
		log = () => {},
	} = options;
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
		room = {
			doc,
			awareness,
			conns: new Map(),
			owners: new Map([[awareness.clientID, SERVER_OWNER]]),
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

	function handleSync(ws, name) {
		const room = getRoom(name);
		room.conns.set(ws, new Set());
		ws.binaryType = 'arraybuffer';
		const bucket = createBucket(limits.messageRate, limits.messageBurst);
		ws.on('message', (data) => {
			if (!bucket.take()) return void ws.close(1008, 'rate limit');
			try {
				const decoder = decoding.createDecoder(new Uint8Array(data));
				const type = decoding.readVarUint(decoder);
				if (type === MESSAGE_SYNC) {
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
				}),
			);
		}
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
		const ip = req.socket.remoteAddress ?? '';
		const open = syncWss.clients.size + signalWss.clients.size;
		const rooms = m[1] === 'sync' ? syncRooms : signalRooms;
		if (
			open >= limits.maxConnections ||
			(perIp.get(ip) ?? 0) >= limits.maxConnectionsPerIp ||
			(!rooms.has(name) && rooms.size >= limits.maxRooms)
		)
			return reject(socket, '503 Service Unavailable');
		perIp.set(ip, (perIp.get(ip) ?? 0) + 1);
		const wss = m[1] === 'sync' ? syncWss : signalWss;
		wss.handleUpgrade(req, socket, head, (ws) => {
			ws.on('close', () => {
				const left = (perIp.get(ip) ?? 1) - 1;
				if (left <= 0) perIp.delete(ip);
				else perIp.set(ip, left);
			});
			try {
				if (m[1] === 'sync') handleSync(ws, name);
				else handleSignal(ws, name);
			} catch (error) {
				log(`${m[1]} ${name}: setup failed (${error.message})`);
				ws.terminate();
			}
		});
	}
	http.on('upgrade', (req, socket, head) => {
		socket.on('error', () => socket.destroy());
		try {
			onUpgrade(req, socket, head);
		} catch (error) {
			log(`upgrade failed (${error.message})`);
			socket.destroy();
		}
	});

	return {
		http,
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
	const server = createTeamsServer({ insecure, log: (m) => console.log(m) });
	const addr = await server.listen();
	const auth = process.env.TEAMS_TOKEN
		? 'token required'
		: 'no token (loopback only unless TEAMS_INSECURE=1)';
	console.log(`teams server on http://${addr.address}:${addr.port}  [${auth}]`);
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
