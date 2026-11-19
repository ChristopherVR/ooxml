import type { VirtualWindow } from './virtualized-list';

/** Which rows of a virtual window have their deferred content. */
export interface DeferredRows {
	/**
	 * Match the rows now in the window: rows new to it wait for a deferred
	 * fill, and rows that left forget theirs, so a row scrolled back into view
	 * waits again rather than building during the scroll.
	 */
	sync(keys: readonly string[]): void;
	/** Drop every pending fill. */
	destroy(): void;
}

/**
 * Track deferred row content for a binding whose rows render declaratively:
 * `onChange` receives the new set of ready row keys after each fill, which the
 * binding stores in its reactive state and checks per row.
 */
export function createDeferredRows(
	view: VirtualWindow,
	onChange: (ready: ReadonlySet<string>) => void,
): DeferredRows {
	let ready = new Set<string>();
	let inWindow = new Set<string>();
	const publish = (next: Set<string>): void => {
		ready = next;
		onChange(next);
	};
	return {
		sync(keys) {
			const next = new Set(keys);
			let pruned = false;
			for (const key of inWindow) {
				if (!next.has(key)) {
					view.cancelDeferred(key);
					pruned ||= ready.has(key);
				}
			}
			if (pruned) {
				publish(new Set([...ready].filter((key) => next.has(key))));
			}
			for (const key of next) {
				if (!inWindow.has(key) && !ready.has(key)) {
					// One row per frame: the binding renders it after this returns.
					view.defer(key, () => {
						if (!inWindow.has(key)) {
							return false;
						}
						publish(new Set(ready).add(key));
						return true;
					});
				}
			}
			inWindow = next;
		},
		destroy() {
			for (const key of inWindow) {
				view.cancelDeferred(key);
			}
			inWindow = new Set();
		},
	};
}
