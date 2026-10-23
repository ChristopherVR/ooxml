/**
 * The headless controller convention every PowerPoint binding adapts.
 *
 * Logic that all five bindings (React, Vue, Svelte, Angular, vanilla) need is
 * written once as a **controller**: a plain object with
 *
 *  - `getSnapshot()`: the current state, an immutable value whose identity
 *    changes only when the state changes (so React's `useSyncExternalStore`
 *    and an `Object.is` check in any other binding can tell "changed" apart);
 *  - `subscribe(listener)`: registers a no-argument listener called after every
 *    change, and returns the function that unregisters it;
 *  - typed actions: ordinary methods that change the state and notify.
 *
 * A controller imports no framework. Each binding ships ONE tiny store adapter
 * that works for every controller (`useController` in React and Vue,
 * `controllerState` in Svelte, `controllerSignal` in Angular, and a direct
 * `subscribe` in vanilla), so a new controller reaches all five bindings
 * without per-binding state code. Derived values that depend on binding state
 * the controller does not own (the live selection, for example) are pure
 * functions next to the controller, which each binding calls from its own
 * reactive scope.
 *
 * @module render/controller/controller
 */

/** A framework-free state owner the bindings adapt (see the module note). */
export interface PptxController<S> {
	/** The current state. Same reference until the state changes. */
	getSnapshot(): S;
	/** Call `listener` after every change; returns the unsubscribe function. */
	subscribe(listener: () => void): () => void;
}

/** The mutable side of a controller's state, kept private to the controller. */
export interface ControllerStore<S> extends PptxController<S> {
	/** Replace the state and notify, unless `next` is the current state. */
	set(next: S): void;
	/** Derive the next state from the current one (see {@link set}). */
	update(recipe: (current: S) => S): void;
}

/**
 * The state cell a controller is built on: holds one immutable snapshot and
 * notifies subscribers when it is replaced by a different value.
 */
export function createControllerStore<S>(initial: S): ControllerStore<S> {
	let snapshot = initial;
	const listeners = new Set<() => void>();
	const set = (next: S): void => {
		if (Object.is(next, snapshot)) {
			return;
		}
		snapshot = next;
		for (const listener of [...listeners]) {
			listener();
		}
	};
	return {
		getSnapshot: () => snapshot,
		subscribe(listener) {
			listeners.add(listener);
			return () => {
				listeners.delete(listener);
			};
		},
		set,
		update: (recipe) => set(recipe(snapshot)),
	};
}
