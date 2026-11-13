/**
 * The host-owned room the building-blocks demo page joins from its panel ("Host join room"):
 * the host owns the Y.Doc, awareness and WebSocket provider and hands the blocks only the public
 * `collaboration` config with an `externalSession`, as `host-owned-collaboration.ts` does for the
 * custom shells. Nothing connects until the button is pressed; the relay comes from `?server=`
 * (a local development relay by default) and the room from `?room=`.
 */

type RoomStatus = 'connecting' | 'connected' | 'disconnected';

export interface BuildingBlocksRoom {
	/** The `collaboration` option to hand the blocks. */
	readonly config: unknown;
	destroy(): void;
}

function relayUrl(params: URLSearchParams): string {
	const url = new URL(params.get('server') ?? 'ws://localhost:1234');
	const loopback = ['localhost', '127.0.0.1', '[::1]'].includes(url.hostname);
	if (url.protocol !== 'wss:' && !(url.protocol === 'ws:' && loopback)) {
		throw new Error('Use a secure WebSocket relay or a local development server.');
	}
	return url.toString().replace(/\/$/u, '');
}

export async function openBuildingBlocksRoom(): Promise<BuildingBlocksRoom> {
	const [{ Awareness }, { WebsocketProvider }, Y] = await Promise.all([
		import('y-protocols/awareness'),
		import('y-websocket'),
		import('yjs'),
	]);
	const params = new URLSearchParams(location.search);
	const doc = new Y.Doc();
	const awareness = new Awareness(doc);
	const provider = new WebsocketProvider(
		relayUrl(params),
		params.get('room') || 'building-blocks-demo',
		doc,
		{ awareness, disableBc: true },
	);
	const listeners = new Set<() => void>();
	let status: RoomStatus = 'connecting';
	const notify = (): void => {
		for (const listener of [...listeners]) {
			listener();
		}
	};
	provider.on('status', (event: { status: string }) => {
		status =
			event.status === 'connected'
				? 'connected'
				: event.status === 'connecting'
					? 'connecting'
					: 'disconnected';
		notify();
	});
	provider.on('sync', notify);
	const config = {
		roomId: params.get('room') || 'building-blocks-demo',
		serverUrl: '',
		userName: params.get('name') || 'Blocks host',
		sessionIntent: params.get('intent') === 'join' ? ('join' as const) : ('create' as const),
		externalSession: {
			doc,
			awareness,
			getSnapshot: () => ({ status, synced: provider.synced }),
			subscribe: (listener: () => void) => {
				listeners.add(listener);
				return () => {
					listeners.delete(listener);
				};
			},
		},
	};
	return {
		config,
		destroy() {
			listeners.clear();
			provider.destroy();
			awareness.destroy();
			doc.destroy();
		},
	};
}
