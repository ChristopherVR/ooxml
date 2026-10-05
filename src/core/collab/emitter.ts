// A tiny typed event emitter: the "typed events" surface of the collab area. Listener errors are
// isolated so one faulty subscriber cannot break the session or the other subscribers.

export type Unsubscribe = () => void;

export class Emitter<Events extends { [K in keyof Events]: unknown }> {
	private readonly listeners = new Map<keyof Events, Set<(payload: never) => void>>();

	on<K extends keyof Events>(event: K, listener: (payload: Events[K]) => void): Unsubscribe {
		let set = this.listeners.get(event);
		if (!set) this.listeners.set(event, (set = new Set()));
		set.add(listener as (payload: never) => void);
		return () => this.off(event, listener);
	}

	off<K extends keyof Events>(event: K, listener: (payload: Events[K]) => void): void {
		this.listeners.get(event)?.delete(listener as (payload: never) => void);
	}

	emit<K extends keyof Events>(event: K, payload: Events[K]): void {
		for (const listener of [...(this.listeners.get(event) ?? [])]) {
			try {
				(listener as (payload: Events[K]) => void)(payload);
			} catch (error) {
				// A throwing subscriber must not break delivery; surface it asynchronously.
				queueMicrotask(() => {
					throw error;
				});
			}
		}
	}

	clear(): void {
		this.listeners.clear();
	}
}
