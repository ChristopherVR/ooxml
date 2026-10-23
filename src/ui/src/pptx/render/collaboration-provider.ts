/**
 * collaboration-provider.ts: the Yjs transport provider (y-websocket or
 * y-webrtc) for a PowerPoint collaboration session, normalised into one handle.
 *
 * Every binding used to carry its own copy of this plumbing. The Yjs packages
 * are optional peer dependencies of the bindings, so this module never imports
 * them: the binding passes {@link CollabTransportModules} loaders (plain
 * `import('y-websocket')` calls in its own package, which keeps the packages
 * tree-shaken when collaboration is unused and resolvable from the binding).
 *
 *  - `websocket`: `new WebsocketProvider(serverUrl, roomId, doc, { params })`;
 *    status events carry `{ status }`; `wsconnected` reports the live state.
 *  - `webrtc`: `new WebrtcProvider(roomId, doc, { signaling, password })`; needs
 *    no document server. Same-browser tabs meet over BroadcastChannel
 *    immediately, so the handle reports `connectedNow: true`.
 *
 * @module render/collaboration-provider
 */
import type { CollaborationConfig, CollaborationTransport } from '../types';
import { createDepartureChannel } from './collaboration-departure';
import type { AwarenessLike } from './collaboration-presence-publisher';
import { clearLocalAwareness } from './collaboration-teardown';

/** A transport-agnostic view of a live Yjs provider. */
export interface CollabProviderHandle {
	/** The provider's awareness instance. */
	awareness: AwarenessLike;
	/**
	 * Subscribe to connection-status changes. The callback receives whether the
	 * transport currently reports a connection.
	 */
	onStatus: (cb: (connected: boolean) => void) => void;
	/** Whether the transport reports a connection immediately after creation. */
	connectedNow: boolean;
	/**
	 * Subscribe to the provider's initial-document-sync confirmation. Websocket
	 * fires reliably once the server sync completes; webrtc only fires when a
	 * peer syncs with us (a lone fresh-room peer never receives one, so callers
	 * pair this with a grace timer, see `createSyncGate`).
	 */
	onSynced: (cb: () => void) => void;
	/** Whether the provider already reports its initial sync as complete. */
	syncedNow: boolean;
	/** Tear the provider down (departure, awareness withdrawal, disconnect, destroy). */
	destroy: () => void;
}

/** The provider surface shared by `y-websocket` and `y-webrtc`. */
interface YProviderLike {
	awareness: unknown;
	on(event: string, cb: (payload: never) => void): void;
	destroy(): void;
}

interface YWebsocketProviderLike extends YProviderLike {
	wsconnected: boolean;
	synced: boolean;
	disconnect(): void;
}

/** `import('y-websocket')`, structurally. */
export interface YWebsocketModuleLike {
	WebsocketProvider: new (
		serverUrl: string,
		roomId: string,
		doc: never,
		opts?: { params?: Record<string, string> },
	) => YWebsocketProviderLike;
}

/** `import('y-webrtc')`, structurally. */
export interface YWebrtcModuleLike {
	WebrtcProvider: new (
		roomId: string,
		doc: never,
		opts?: { signaling?: string[]; password?: string },
	) => YProviderLike;
}

/** The binding's lazy loaders for the optional transport packages. */
export interface CollabTransportModules {
	loadWebsocket(): Promise<unknown>;
	loadWebrtc(): Promise<unknown>;
}

/**
 * Create and return a normalised provider handle for the requested transport.
 * Callers must destroy any previous provider first: y-webrtc throws when the
 * same room is joined twice within one page.
 */
export async function createCollabProvider(
	transport: CollaborationTransport,
	config: CollaborationConfig,
	doc: unknown,
	modules: CollabTransportModules,
): Promise<CollabProviderHandle> {
	if (transport === 'webrtc') {
		const { WebrtcProvider } = (await modules.loadWebrtc()) as YWebrtcModuleLike;
		const provider = new WebrtcProvider(config.roomId, doc as never, {
			...(config.signaling?.length ? { signaling: config.signaling } : {}),
			...(config.authToken ? { password: config.authToken } : {}),
		});
		const awareness = provider.awareness as AwarenessLike;
		const departure = createDepartureChannel(config.roomId, awareness as never);
		return {
			awareness,
			onStatus: (cb) =>
				provider.on('status', ((event: { connected?: boolean }) =>
					cb(Boolean(event?.connected))) as never),
			connectedNow: true,
			onSynced: (cb) =>
				provider.on('synced', ((event: { synced?: boolean }) => {
					if (event?.synced !== false) {
						cb();
					}
				}) as never),
			syncedNow: false,
			destroy: () => {
				// Announce first: it is synchronous, so it still reaches same-browser
				// peers when this runs from a document that is being destroyed. The
				// provider's own awareness removal is broadcast a microtask later and
				// would be dropped, leaving a ghost collaborator until the 30s
				// awareness timeout.
				departure.announce();
				departure.dispose();
				clearLocalAwareness(awareness as never);
				provider.destroy();
			},
		};
	}

	const { WebsocketProvider } = (await modules.loadWebsocket()) as YWebsocketModuleLike;
	const provider = new WebsocketProvider(
		config.serverUrl,
		config.roomId,
		doc as never,
		config.authToken ? { params: { token: config.authToken } } : undefined,
	);
	const awareness = provider.awareness as AwarenessLike;
	const departure = createDepartureChannel(config.roomId, awareness as never);
	return {
		awareness,
		onStatus: (cb) =>
			provider.on('status', ((event: { status?: string }) => {
				if (event?.status === 'connected') {
					cb(true);
				} else if (event?.status === 'disconnected') {
					cb(false);
				}
			}) as never),
		connectedNow: provider.wsconnected,
		onSynced: (cb) =>
			provider.on('sync', ((isSynced: boolean) => {
				if (isSynced) {
					cb();
				}
			}) as never),
		syncedNow: provider.synced,
		destroy: () => {
			departure.announce();
			departure.dispose();
			clearLocalAwareness(awareness as never);
			provider.disconnect();
			provider.destroy();
		},
	};
}
