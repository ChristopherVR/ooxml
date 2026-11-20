// Test-only: what an element holds on to outside itself, so the conformance suite can tell that a
// removed element let go of it. Counts the listeners on `window` and `document`, the observers
// still observing and the intervals still running.

type Release = () => void;

export interface ResourceTracker {
	/** The registrations still live, as `kind:detail` strings (a multiset, so repeats count). */
	live(): string[];
	/** Puts back every wrapped method. */
	restore(): void;
}

const OBSERVERS = ['ResizeObserver', 'MutationObserver', 'IntersectionObserver'] as const;

export function trackResources(win: Window & typeof globalThis): ResourceTracker {
	const live = new Map<object, string>();
	const restore: Release[] = [];
	const wrap = <T extends object, K extends keyof T>(
		target: T,
		key: K,
		make: (old: T[K]) => T[K],
	) => {
		const old = target[key];
		target[key] = make(old);
		restore.push(() => {
			target[key] = old;
		});
	};

	for (const [name, target] of [
		['window', win],
		['document', win.document],
	] as const) {
		const keyOf = (type: string, options?: unknown): string => {
			const capture =
				typeof options === 'boolean'
					? options
					: Boolean((options as EventListenerOptions)?.capture);
			return `${name}:${type}:${capture}`;
		};
		const tokens = new Map<unknown, Map<string, object>>();
		wrap(
			target,
			'addEventListener',
			(add) =>
				function (this: EventTarget, type: string, listener: unknown, options?: unknown) {
					if (
						listener &&
						!(typeof options === 'object' && (options as AddEventListenerOptions)?.once)
					) {
						const key = keyOf(type, options);
						const byListener = tokens.get(listener) ?? new Map<string, object>();
						if (!byListener.has(key)) {
							const handle = {};
							byListener.set(key, handle);
							tokens.set(listener, byListener);
							live.set(handle, `listener ${key}`);
						}
					}
					return (add as (...args: unknown[]) => void).call(this, type, listener, options);
				} as typeof add,
		);
		wrap(
			target,
			'removeEventListener',
			(remove) =>
				function (this: EventTarget, type: string, listener: unknown, options?: unknown) {
					const key = keyOf(type, options);
					const handle = tokens.get(listener)?.get(key);
					if (handle) {
						live.delete(handle);
						tokens.get(listener)!.delete(key);
					}
					return (remove as (...args: unknown[]) => void).call(this, type, listener, options);
				} as typeof remove,
		);
	}

	for (const name of OBSERVERS) {
		const Observer = win[name] as
			| { prototype: { observe: unknown; disconnect: unknown } }
			| undefined;
		if (!Observer) continue;
		const proto = Observer.prototype as {
			observe: (...args: unknown[]) => void;
			disconnect: () => void;
		};
		wrap(
			proto,
			'observe',
			(observe) =>
				function (this: object, ...args: unknown[]) {
					live.set(this, `observer ${name}`);
					return observe.apply(this, args);
				},
		);
		wrap(
			proto,
			'disconnect',
			(disconnect) =>
				function (this: object) {
					live.delete(this);
					return disconnect.call(this);
				},
		);
	}

	const intervals = new Map<unknown, object>();
	wrap(
		win,
		'setInterval',
		(set) =>
			((...args: Parameters<typeof setInterval>) => {
				const handle = set(...args);
				const token = {};
				intervals.set(handle, token);
				live.set(token, 'interval');
				return handle;
			}) as typeof set,
	);
	wrap(
		win,
		'clearInterval',
		(clear) =>
			((handle?: Parameters<typeof clearInterval>[0]) => {
				const token = intervals.get(handle);
				if (token) live.delete(token);
				intervals.delete(handle);
				clear(handle);
			}) as typeof clear,
	);

	return {
		live: () => [...live.values()].sort(),
		restore: () => {
			for (const undo of restore.splice(0).reverse()) undo();
		},
	};
}

/** The entries of `after` that `before` does not account for (multiset difference). */
export function leftOver(before: readonly string[], after: readonly string[]): string[] {
	const counts = new Map<string, number>();
	for (const entry of before) counts.set(entry, (counts.get(entry) ?? 0) + 1);
	const extra: string[] = [];
	for (const entry of after) {
		const left = counts.get(entry) ?? 0;
		if (left > 0) counts.set(entry, left - 1);
		else extra.push(entry);
	}
	return extra;
}
