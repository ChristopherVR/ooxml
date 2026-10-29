// The sync transport a server config selects for one room: the configured y-websocket server, or
// same-browser tabs over BroadcastChannel when there is none. Shared by the team workspace and the
// shared assistant chats, so both reach the same server the same way. DOM-free. New code.
import { createBroadcastTransport } from '../collab/broadcast-transport';
import type { BroadcastChannelConstructor } from '../collab/broadcast-transport';
import type { Transport } from '../collab/provider';
import { createWebSocketTransport } from '../collab/websocket-transport';
import type { WebSocketConstructor } from '../collab/websocket-transport';
import { protocolParams } from './protocol';
import { type TeamsServerConfig, connectionAuth, endpointUrl } from './server-config';

export interface SyncTransportClasses {
	WebSocket?: WebSocketConstructor | undefined;
	BroadcastChannel?: BroadcastChannelConstructor | undefined;
}

export function createSyncTransport(
	config: TeamsServerConfig,
	roomId: string,
	classes: SyncTransportClasses = {},
	/** Extra query parameters for the server, such as who opens a private chat room. */
	query?: Record<string, string>,
): Transport {
	// With accounts the server reads the user from the signed session, not from a claimed id.
	const auth = connectionAuth(config);
	const params = { ...query, ...protocolParams(), ...auth.params };
	return config.mode === 'server' && config.syncUrl
		? createWebSocketTransport({
				url: endpointUrl(config.syncUrl, roomId, auth.token, params),
				protocols: auth.protocols,
				...(classes.WebSocket ? { WebSocket: classes.WebSocket } : {}),
			})
		: createBroadcastTransport({
				roomId,
				...(classes.BroadcastChannel ? { BroadcastChannel: classes.BroadcastChannel } : {}),
			});
}
