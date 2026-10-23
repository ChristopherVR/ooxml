import type { Equality } from './equality';

/**
 * Remembers the last value forwarded for each prop. Calling it returns true the first time and then
 * only when the prop itself changed, so a parent re-render never undoes a value the user changed
 * inside the element (Editing / Viewing, File > Open renaming the document, Customize Ribbon).
 */
export interface PropTracker<K extends string = string> {
	<T>(key: K, value: T, same?: Equality<T>): boolean;
	/** Records `value` as already forwarded, so an equal first value is not written. */
	seed(key: K, value: unknown): void;
	/** The value last forwarded for `key`; `undefined` before the first. */
	sent(key: K): unknown;
}

export function createPropTracker<K extends string = string>(): PropTracker<K> {
	const sent = new Map<K, unknown>();
	const changed = <T>(key: K, value: T, same: Equality<T> = Object.is): boolean => {
		if (sent.has(key) && same(sent.get(key) as T, value)) return false;
		sent.set(key, value);
		return true;
	};
	return Object.assign(changed, {
		seed: (key: K, value: unknown) => void sent.set(key, value),
		sent: (key: K) => sent.get(key),
	});
}
