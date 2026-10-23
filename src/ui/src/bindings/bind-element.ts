// The element wiring every framework adapter repeats, written once: listen to the contract's events
// and call the current handler options, and forward changed props by each prop's policy.
import type { BindingContract, EventRule, PropRule } from './contract';
import { createPropTracker } from './tracker';

/** Stops some (or, without names, all) of the listeners `listenEvents` added. */
export type Unlisten = (names?: Iterable<string>) => void;

/**
 * Listens to every event in `events` on `element`; each one calls the handler option the rule
 * names, read through `handlers()` at event time, with the rule's arguments. `onEvent` runs first
 * (a product records what the element emitted, so it is not echoed back as a prop).
 */
export function listenEvents(
	element: EventTarget,
	events: Readonly<Record<string, EventRule>>,
	handlers: () => object,
	onEvent?: (name: string, event: CustomEvent) => void,
): Unlisten {
	const listeners = new Map<string, EventListener>();
	for (const [name, rule] of Object.entries(events)) {
		const listener = (event: Event) => {
			const custom = event as CustomEvent;
			onEvent?.(name, custom);
			const handler = (handlers() as Record<string, unknown>)[rule.option];
			if (typeof handler === 'function')
				(handler as (...args: unknown[]) => void)(...(rule.args?.(custom) ?? [custom.detail]));
		};
		listeners.set(name, listener);
		element.addEventListener(name, listener);
	}
	return (names = [...listeners.keys()]) => {
		for (const name of names) {
			const listener = listeners.get(name);
			if (!listener) continue;
			element.removeEventListener(name, listener);
			listeners.delete(name);
		}
	};
}

export interface BoundElement<P> {
	/** Forwards the props whose policy says so (every policy but `manual`). */
	update(next: { readonly [K in keyof P]?: P[K] | undefined }): void;
	/** The value last written for `key`. */
	sent(key: keyof P & string): unknown;
	unlisten: Unlisten;
}

export interface BindElementOptions {
	onEvent?: (name: string, event: CustomEvent) => void;
}

/**
 * Wires `element` to a product contract. `handlers()` returns the current handler options (the
 * latest props), read at event time. Product-specific steps (a document model that must follow
 * collaboration, `src` supersession) stay in the product's mount function around `update`.
 */
export function bindElement<E extends EventTarget, P extends object>(
	element: E,
	contract: Pick<BindingContract<E, P>, 'props' | 'events'>,
	handlers: () => object,
	options: BindElementOptions = {},
): BoundElement<P> {
	type Key = keyof P & string;
	const tracker = createPropTracker<Key>();
	const rules = Object.entries(contract.props) as [Key, PropRule<E, unknown>][];
	for (const [key, rule] of rules) if ('initial' in rule) tracker.seed(key, rule.initial);
	const unlisten = listenEvents(element, contract.events, handlers, options.onEvent);
	return {
		update(next) {
			const values = next as Record<string, unknown>;
			for (const [key, rule] of rules) {
				if (rule.forward === 'manual') continue;
				const raw = values[key];
				if (rule.forward === 'defined' && raw === undefined) continue;
				const value = raw === undefined && rule.fallback !== undefined ? rule.fallback : raw;
				if (rule.forward !== 'always' && !tracker(key, value, rule.same)) continue;
				if (rule.apply) rule.apply(element, value);
				else (element as unknown as Record<string, unknown>)[key] = value;
			}
		},
		sent: (key) => tracker.sent(key),
		unlisten,
	};
}

const tokens = (value: string | undefined): string[] => value?.split(/\s+/u).filter(Boolean) ?? [];

/** Swaps the classes a binding added (`previous`) for `next`, leaving the element's own classes. */
export function syncClassName(
	element: Element,
	previous: string | undefined,
	next: string | undefined,
): void {
	if (previous === next) return;
	element.classList.remove(...tokens(previous));
	element.classList.add(...tokens(next));
}
