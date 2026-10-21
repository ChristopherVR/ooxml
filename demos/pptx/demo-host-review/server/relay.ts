// The collaboration relay: `openteams-server`'s y-websocket sync rooms, served from the hub's own
// origin, with the host deciding every socket through `authorizeSync`. A room ticket admits a
// socket for writing or read-only (the server drops a read-only socket's document updates); the
// hub closes an account's sockets when its access changes, and they reconnect with a new ticket.
import { mkdtempSync, rmSync } from 'node:fs';
import type { IncomingMessage } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { Duplex } from 'node:stream';

import { createTeamsServer } from 'openteams-server';

import { canWrite } from './directory';
import type { Ticket } from './tickets';

/** Close code a client sees when the host changed or removed its access. */
export const ACCESS_CHANGED = 4403;
const SEED_LEASE_MS = 60_000;

export function createRelay(options: { verify: (token: string) => Ticket | null }) {
	// Rooms persist to disk while the process runs; a real host keeps this directory.
	const dataDir = mkdtempSync(join(tmpdir(), 'review-hub-'));
	const teams = createTeamsServer({
		dataDir,
		authorizeSync({ room, url }) {
			const ticket = options.verify(url.searchParams.get('ticket') ?? '');
			if (!ticket || ticket.review !== room) return null;
			return { access: canWrite(ticket.permission) ? 'write' : 'read', principal: ticket.account };
		},
	});
	const leases = new Map<string, { account: string; until: number }>();

	return {
		/** Handle `/sync/<review>?ticket=` upgrades; returns false for other paths (Vite HMR). */
		upgrade(request: IncomingMessage, socket: Duplex, head: Buffer): boolean {
			if (!request.url?.startsWith('/sync/')) return false;
			teams.handleUpgrade(request, socket, head);
			return true;
		},
		/**
		 * Whether `account` should seed an empty room. The host, not the clients, decides who creates
		 * it: one writer at a time holds a short lease, so two first arrivals never both publish their
		 * startup deck.
		 */
		claimSeed(room: string, account: string, now = Date.now()): boolean {
			if (!teams.isSyncRoomEmpty(room)) return false;
			const lease = leases.get(room);
			if (lease && lease.until > now && lease.account !== account) return false;
			leases.set(room, { account, until: now + SEED_LEASE_MS });
			return true;
		},
		/** Drop an account's sockets after its access changed. */
		revoke(room: string, account: string): void {
			teams.closeSync(room, account, ACCESS_CHANGED, 'access changed');
		},
		/** The accounts with an open socket in the room. */
		online: (room: string): string[] => teams.syncPrincipals(room),
		async close(): Promise<void> {
			await teams.close();
			rmSync(dataDir, { recursive: true, force: true });
		},
	};
}

export type Relay = ReturnType<typeof createRelay>;
