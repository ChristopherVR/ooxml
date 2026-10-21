// A y-websocket compatible relay that admits a socket only with a host ticket and enforces the
// ticket's access on every message. Rooms live as long as the process, so a review keeps its
// shared state when everybody leaves; a real host persists the Yjs updates as well (see README).
import type { IncomingMessage } from 'node:http';
import type { Duplex } from 'node:stream';

import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import { WebSocket, WebSocketServer } from 'ws';
import * as awarenessProtocol from 'y-protocols/awareness';
import * as syncProtocol from 'y-protocols/sync';
import * as Y from 'yjs';

import { canWrite } from './directory';
import type { Permission } from './directory';
import {
	MESSAGE_AWARENESS,
	MESSAGE_QUERY_AWARENESS,
	MESSAGE_SYNC,
	mayProcess,
	messageKind,
	ownAwareness,
} from './relay-policy';
import type { Ticket } from './tickets';

/** Close code a client sees when the host changed or removed its access. */
export const ACCESS_CHANGED = 4403;
export const MAX_MESSAGE_BYTES = 32 * 1024 * 1024;
const SEED_LEASE_MS = 60_000;

interface Connection {
	socket: WebSocket;
	account: string;
	permission: Permission;
}

interface Room {
	doc: Y.Doc;
	awareness: awarenessProtocol.Awareness;
	connections: Set<Connection>;
	owners: Map<number, Connection>;
	seedLease?: { account: string; until: number };
}

const frame = (type: number, write: (encoder: encoding.Encoder) => void): Uint8Array => {
	const encoder = encoding.createEncoder();
	encoding.writeVarUint(encoder, type);
	write(encoder);
	return encoding.toUint8Array(encoder);
};

const send = (connection: Connection, bytes: Uint8Array): void => {
	if (connection.socket.readyState === WebSocket.OPEN) connection.socket.send(bytes);
};

export function createRelay(options: {
	verify: (token: string) => Ticket | null;
	onPresence?: (review: string, accounts: string[]) => void;
}) {
	const rooms = new Map<string, Room>();
	const server = new WebSocketServer({ noServer: true, maxPayload: MAX_MESSAGE_BYTES });

	const room = (name: string): Room => {
		const existing = rooms.get(name);
		if (existing) return existing;
		const doc = new Y.Doc();
		const awareness = new awarenessProtocol.Awareness(doc);
		awareness.setLocalState(null);
		const created: Room = { doc, awareness, connections: new Set(), owners: new Map() };
		doc.on('update', (update: Uint8Array, origin: unknown) => {
			const message = frame(MESSAGE_SYNC, (e) => syncProtocol.writeUpdate(e, update));
			for (const connection of created.connections)
				if (connection !== origin) send(connection, message);
		});
		awareness.on(
			'update',
			({ added, updated, removed }: Record<'added' | 'updated' | 'removed', number[]>) => {
				const changed = [...added, ...updated, ...removed];
				const message = frame(MESSAGE_AWARENESS, (e) =>
					encoding.writeVarUint8Array(
						e,
						awarenessProtocol.encodeAwarenessUpdate(awareness, changed),
					),
				);
				for (const connection of created.connections) send(connection, message);
			},
		);
		rooms.set(name, created);
		return created;
	};

	const online = (name: string, target: Room): void =>
		options.onPresence?.(name, [...new Set([...target.connections].map((c) => c.account))]);

	const handle = (target: Room, connection: Connection, bytes: Uint8Array): void => {
		const kind = messageKind(bytes);
		if (!mayProcess(kind, canWrite(connection.permission))) return;
		const decoder = decoding.createDecoder(bytes);
		const type = decoding.readVarUint(decoder);
		if (type === MESSAGE_SYNC) {
			const encoder = encoding.createEncoder();
			encoding.writeVarUint(encoder, MESSAGE_SYNC);
			syncProtocol.readSyncMessage(decoder, encoder, target.doc, connection);
			if (encoding.length(encoder) > 1) send(connection, encoding.toUint8Array(encoder));
		} else if (type === MESSAGE_AWARENESS) {
			const update = ownAwareness(decoding.readVarUint8Array(decoder), connection, target.owners);
			if (update) awarenessProtocol.applyAwarenessUpdate(target.awareness, update, connection);
		} else if (type === MESSAGE_QUERY_AWARENESS) {
			const ids = [...target.awareness.getStates().keys()];
			send(
				connection,
				frame(MESSAGE_AWARENESS, (e) =>
					encoding.writeVarUint8Array(
						e,
						awarenessProtocol.encodeAwarenessUpdate(target.awareness, ids),
					),
				),
			);
		}
	};

	const join = (name: string, socket: WebSocket, ticket: Ticket): void => {
		const target = room(name);
		const connection: Connection = {
			socket,
			account: ticket.account,
			permission: ticket.permission,
		};
		target.connections.add(connection);
		socket.binaryType = 'nodebuffer';
		socket.on('message', (data: Buffer) => {
			try {
				handle(target, connection, new Uint8Array(data.buffer, data.byteOffset, data.byteLength));
			} catch {
				socket.close(1003, 'malformed message');
			}
		});
		socket.on('close', () => {
			target.connections.delete(connection);
			const ids = [...target.owners].filter(([, owner]) => owner === connection).map(([id]) => id);
			for (const id of ids) target.owners.delete(id);
			awarenessProtocol.removeAwarenessStates(target.awareness, ids, null);
			online(name, target);
		});
		send(
			connection,
			frame(MESSAGE_SYNC, (e) => syncProtocol.writeSyncStep1(e, target.doc)),
		);
		const ids = [...target.awareness.getStates().keys()];
		if (ids.length > 0) {
			send(
				connection,
				frame(MESSAGE_AWARENESS, (e) =>
					encoding.writeVarUint8Array(
						e,
						awarenessProtocol.encodeAwarenessUpdate(target.awareness, ids),
					),
				),
			);
		}
		online(name, target);
	};

	return {
		/** Handle an HTTP upgrade for `/collab/<review>?ticket=`; returns false for other paths. */
		upgrade(request: IncomingMessage, socket: Duplex, head: Buffer): boolean {
			const url = new URL(request.url ?? '/', 'http://relay');
			if (!url.pathname.startsWith('/collab/')) return false;
			const name = decodeURIComponent(url.pathname.slice('/collab/'.length));
			const ticket = options.verify(url.searchParams.get('ticket') ?? '');
			if (!ticket || ticket.review !== name || ticket.permission === 'none') {
				socket.end('HTTP/1.1 401 Unauthorized\r\nConnection: close\r\n\r\n');
				return true;
			}
			server.handleUpgrade(request, socket, head, (ws) => join(name, ws, ticket));
			return true;
		},
		/**
		 * Whether `account` should seed an empty room. The host, not the clients, decides who
		 * creates it: one writer at a time holds a short lease, so two first arrivals never both
		 * publish their startup deck.
		 */
		claimSeed(name: string, account: string, now = Date.now()): boolean {
			const target = room(name);
			if (target.doc.store.clients.size > 0) return false;
			const lease = target.seedLease;
			if (lease && lease.until > now && lease.account !== account) return false;
			target.seedLease = { account, until: now + SEED_LEASE_MS };
			return true;
		},
		/** Drop an account's sockets after its access changed; the client reconnects with a new ticket. */
		revoke(name: string, account: string): void {
			for (const connection of rooms.get(name)?.connections ?? []) {
				if (connection.account === account)
					connection.socket.close(ACCESS_CHANGED, 'access changed');
			}
		},
		close(): void {
			for (const target of rooms.values()) {
				for (const connection of target.connections) connection.socket.terminate();
				target.awareness.destroy();
				target.doc.destroy();
			}
			rooms.clear();
			server.close();
		},
	};
}

export type Relay = ReturnType<typeof createRelay>;
