// In-memory mesh transport: every transport created on the same hub and room hears every other.
// It is the reference implementation of {@link Transport} for tests, demos, offline/same-process
// co-editing and local-first prototypes. Delivery is synchronous by default and can be made
// asynchronous (one microtask) to exercise ordering; `link` simulates a network partition.
// New code written for the collab area.
import type { Transport, TransportHandlers } from './provider.js';

export interface MemoryTransport extends Transport {
	/** Whether this endpoint is currently attached to its room. */
	readonly connected: boolean;
}

export interface MemoryHub {
	/** Create an endpoint for `roomId`; it hears the other endpoints of the same room. */
	createTransport: (roomId: string) => MemoryTransport;
	/** Number of connected endpoints in a room. */
	size: (roomId: string) => number;
}

export interface MemoryHubOptions {
	/** Deliver through `queueMicrotask` instead of synchronously. */
	async?: boolean;
	/** Return false to drop a message (simulate loss); receives raw bytes and the sender index. */
	filter?: (data: Uint8Array, from: number) => boolean;
}

interface Endpoint {
	id: number;
	handlers: TransportHandlers;
}

export function createMemoryHub(options: MemoryHubOptions = {}): MemoryHub {
	const rooms = new Map<string, Set<Endpoint>>();
	let nextId = 0;

	const deliver = (task: () => void): void => {
		if (options.async) queueMicrotask(task);
		else task();
	};

	return {
		size: (roomId) => rooms.get(roomId)?.size ?? 0,
		createTransport(roomId) {
			const id = nextId++;
			let endpoint: Endpoint | null = null;
			const peers = (): Set<Endpoint> => {
				let room = rooms.get(roomId);
				if (!room) rooms.set(roomId, (room = new Set()));
				return room;
			};
			return {
				get connected() {
					return endpoint !== null;
				},
				connect(handlers) {
					if (endpoint) peers().delete(endpoint);
					const self: Endpoint = { id, handlers };
					const others = [...peers()];
					peers().add(self);
					endpoint = self;
					deliver(() => {
						if (endpoint !== self) return;
						handlers.open();
						for (const other of others) other.handlers.peer?.();
					});
				},
				send(data) {
					const self = endpoint;
					if (!self || (options.filter && !options.filter(data, id))) return;
					// Copy per receiver: a consumer must never see another consumer's mutations.
					for (const other of [...peers()])
						if (other !== self) deliver(() => other.handlers.message(data.slice()));
				},
				disconnect() {
					const self = endpoint;
					if (!self) return;
					endpoint = null;
					peers().delete(self);
					self.handlers.close('disconnected');
				},
			};
		},
	};
}
