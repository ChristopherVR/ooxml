// The part of `openteams-server` (plain JavaScript) the relay uses.
declare module 'openteams-server' {
	import type { IncomingMessage, Server } from 'node:http';
	import type { Duplex } from 'node:stream';

	type SyncAccess = 'write' | 'read';
	type SyncVerdict =
		| SyncAccess
		| { access: SyncAccess; principal?: string }
		| null
		| undefined
		| false;

	export interface TeamsServer {
		http: Server;
		handleUpgrade(request: IncomingMessage, socket: Duplex, head: Buffer): void;
		closeSync(room: string, principal?: string, code?: number, reason?: string): number;
		syncPrincipals(room: string): string[];
		isSyncRoomEmpty(room: string): boolean;
		close(): Promise<void>;
	}

	export function createTeamsServer(options?: {
		dataDir?: string;
		authorizeSync?: (request: {
			room: string;
			user: string | null;
			url: URL;
			headers: IncomingMessage['headers'];
		}) => SyncVerdict | Promise<SyncVerdict>;
	}): TeamsServer;
}
