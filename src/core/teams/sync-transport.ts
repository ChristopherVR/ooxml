// The sync transport a server config selects for one room: the configured y-websocket server, or
// same-browser tabs over BroadcastChannel when there is none. Shared by the team workspace and the
// shared assistant chats, so both reach the same server the same way. DOM-free. New code.
import { createBroadcastTransport } from '../collab/broadcast-transport';
import type { BroadcastChannelConstructor } from '../collab/broadcast-transport';
import type { Transport } from '../collab/provider';
import { createWebSocketTransport } from '../collab/websocket-transport';
import type { WebSocketConstructor, WebSocketTarget } from '../collab/websocket-transport';
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
	/**
	 * Resolves the token before every connection attempt, in place of `config.token`, for a host
	 * that issues short-lived credentials. A rejection is reported and retried like a refused socket.
	 */
	token?: () => Promise<string>,
): Transport {
	const { syncUrl } = config;
	if (config.mode !== 'server' || !syncUrl)
		return createBroadcastTransport({
			roomId,
			...(classes.BroadcastChannel ? { BroadcastChannel: classes.BroadcastChannel } : {}),
		});
	// With accounts the server reads the user from the signed session, not from a claimed id.
	const target = (current: TeamsServerConfig): WebSocketTarget => {
		const auth = connectionAuth(current);
		const params = { ...query, ...protocolParams(), ...auth.params };
		return { url: endpointUrl(syncUrl, roomId, auth.token, params), protocols: auth.protocols };
	};
	const fixed = target(config);
	return createWebSocketTransport({
		url: token ? async () => target({ ...config, token: await token() }) : fixed.url,
		protocols: fixed.protocols,
		...(classes.WebSocket ? { WebSocket: classes.WebSocket } : {}),
	});
}
