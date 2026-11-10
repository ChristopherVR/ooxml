/**
 * A tiny framework-free reactive store: `get` / `set(patch)` / `subscribe`.
 *
 * This is deliberately minimal (a few dozen lines, no external deps). It only
 * carries the vanilla binding's *view state*; all domain logic (parsing,
 * styles, geometry, text building) stays in `pptx-viewer-core` and
 * `pptx-viewer-shared`.
 */

/** Listener invoked after every state change with the next and previous state. */
export type StoreListener<T> = (state: T, previous: T) => void;

export interface Store<T extends object> {
	/** Current state snapshot (do not mutate). */
	get(): T;
	/** Shallow-merge a partial patch into the state and notify subscribers. */
	set(patch: Partial<T>): void;
	/** Subscribe to state changes; returns an unsubscribe function. */
	subscribe(listener: StoreListener<T>): () => void;
}

/**
 * Delivery passes before a notification gives up: each pass is one round of
 * listeners setting state again. Real chains settle in two or three; a listener
 * that never settles threw a stack overflow while delivery was nested, and
 * throws this instead of spinning.
 */
const MAX_NOTIFY_PASSES = 100;

export function createStore<T extends object>(initial: T): Store<T> {
	let state = initial;
	/** Each listener, with the state it was last told about. */
	const listeners = new Map<StoreListener<T>, T>();
	let notifying = false;

	/**
	 * Tell every listener about the state it has not seen yet. A `set` from inside
	 * a listener is delivered once that listener returns, not nested inside it:
	 * nested, the listeners after it were then called again with the outer
	 * `previous`, so they saw the inner change twice (and rendered twice). Here a
	 * listener not reached yet sees both changes as one, and one already called
	 * sees the inner change on its own, in order. It follows that a listener
	 * unsubscribed by an earlier one is not called for that change, and a writer
	 * reads what later listeners derive from its write only after it returns.
	 */
	const notify = (): void => {
		notifying = true;
		try {
			let stale = true;
			for (let pass = 1; stale; pass++) {
				if (pass > MAX_NOTIFY_PASSES) {
					throw new Error('Store listeners kept changing the state: one never settles.');
				}
				stale = false;
				for (const listener of Array.from(listeners.keys())) {
					const previous = listeners.get(listener);
					if (previous === undefined || previous === state) {
						continue;
					}
					const next = state;
					listeners.set(listener, next);
					listener(next, previous);
					stale ||= state !== next;
				}
			}
		} finally {
			notifying = false;
		}
	};

	return {
		get: () => state,
		set: (patch) => {
			// Skip notification when the patch changes nothing (identity compare
			// per key keeps this O(patch) and predictable).
			let changed = false;
			for (const key of Object.keys(patch) as Array<keyof T>) {
				if (patch[key] !== state[key]) {
					changed = true;
					break;
				}
			}
			if (!changed) {
				return;
			}
			state = { ...state, ...patch };
			if (!notifying) {
				notify();
			}
		},
		subscribe: (listener) => {
			if (!listeners.has(listener)) {
				listeners.set(listener, state);
			}
			return () => {
				listeners.delete(listener);
			};
		},
	};
}
