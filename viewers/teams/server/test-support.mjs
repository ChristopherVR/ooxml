// Helpers shared by the server tests. Not published.
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import { WebSocket } from 'ws';
import * as syncProtocol from 'y-protocols/sync';
import * as Y from 'yjs';

export const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

export async function until(fn, ms = 4000) {
	const end = Date.now() + ms;
	while (Date.now() < end) {
		if (await fn()) return true;
		await sleep(20);
	}
	return false;
}

const protocolsOf = (session) => [
	'openteams',
	...(session ? [`openteams.session.${session}`] : []),
];

const frame = (write) => {
	const encoder = encoding.createEncoder();
	encoding.writeVarUint(encoder, 0);
	write(encoder);
	return encoding.toUint8Array(encoder);
};

/**
 * A y-websocket client of `room` as the holder of `session`: it answers the server's sync step 1,
 * applies what arrives and sends every local change. `status` is `open`, `http-<code>` or `error`;
 * `closed` is `{ code, reason }` once the socket closed.
 */
export function syncClient(port, session, room, { doc = new Y.Doc(), query = '' } = {}) {
	const ws = new WebSocket(`ws://127.0.0.1:${port}/sync/${room}${query}`, protocolsOf(session), {
		maxPayload: 256 * 1024 * 1024,
	});
	ws.binaryType = 'arraybuffer';
	const client = { ws, doc, status: 'connecting', closed: null, messages: [] };
	const send = (bytes) => ws.readyState === 1 && ws.send(bytes);
	client.sendUpdate = (update) => send(frame((e) => syncProtocol.writeUpdate(e, update)));
	doc.on('update', (update, origin) => origin !== client && client.sendUpdate(update));
	client.ready = new Promise((ok) => {
		ws.on('open', () => {
			client.status = 'open';
			send(frame((e) => syncProtocol.writeSyncStep1(e, doc)));
			ok(client);
		});
		ws.on('unexpected-response', (_req, res) => {
			client.status = `http-${res.statusCode}`;
			ok(client);
		});
		ws.on('error', () => {
			if (client.status === 'connecting') client.status = 'error';
			ok(client);
		});
	});
	ws.on('close', (code, reason) => (client.closed = { code, reason: String(reason) }));
	ws.on('message', (data) => {
		const bytes = new Uint8Array(data);
		client.messages.push(bytes);
		const decoder = decoding.createDecoder(bytes);
		if (decoding.readVarUint(decoder) !== 0) return;
		const kind = decoding.readVarUint(decoder);
		if (kind === syncProtocol.messageYjsSyncStep1) {
			const vector = decoding.readVarUint8Array(decoder);
			send(frame((e) => syncProtocol.writeSyncStep2(e, doc, vector)));
		} else Y.applyUpdate(doc, decoding.readVarUint8Array(decoder), client);
	});
	client.close = () => ws.close();
	return client;
}

/**
 * Give the sync room `room` some content as the holder of `session`, as a client opening a
 * workspace does, and resolve once the server has applied it. The server answers messages in order,
 * so its reply to the sync step 1 sent after the update proves the update was handled.
 */
export async function seedRoom(
	port,
	session,
	room,
	write = (doc) => doc.getMap('seed').set('at', 1),
) {
	const ws = new WebSocket(`ws://127.0.0.1:${port}/sync/${room}`, protocolsOf(session), {
		maxPayload: 256 * 1024 * 1024,
	});
	ws.binaryType = 'arraybuffer';
	await new Promise((ok, fail) => (ws.on('open', ok), ws.on('error', fail)));
	const doc = new Y.Doc();
	write(doc);
	const answered = new Promise((ok) => {
		ws.on('close', (code, reason) => ok({ code, reason: String(reason) }));
		ws.on('message', (data) => {
			const decoder = decoding.createDecoder(new Uint8Array(data));
			if (decoding.readVarUint(decoder) !== 0) return;
			if (decoding.readVarUint(decoder) === syncProtocol.messageYjsSyncStep2) ok(null);
		});
	});
	ws.send(frame((e) => syncProtocol.writeUpdate(e, Y.encodeStateAsUpdate(doc))));
	ws.send(frame((e) => syncProtocol.writeSyncStep1(e, new Y.Doc())));
	const closed = await answered;
	ws.close();
	return closed;
}
