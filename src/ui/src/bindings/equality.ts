// Change detection for forwarded props: a framework passes a fresh array or object literal on
// every render, so "changed" has to mean "different by value" for those props.

/** Compares two values the way a prop forwarder does. */
export type Equality<T = unknown> = (a: T, b: T) => boolean;

/** Same items in the same order; `undefined` and `null` count as an empty list. */
export function sameList(
	a: readonly unknown[] | undefined | null,
	b: readonly unknown[] | undefined | null,
): boolean {
	if (a === b) return true;
	const left = a ?? [];
	const right = b ?? [];
	return left.length === right.length && left.every((item, index) => item === right[index]);
}

/**
 * Shallow, by value: an inline object (`themeColors`) passed on every render is not a change.
 * `undefined` and `null` count as an empty object.
 */
export function shallowEqual(a: object | undefined | null, b: object | undefined | null): boolean {
	if (a === b) return true;
	const left: Record<string, unknown> = { ...a };
	const right: Record<string, unknown> = { ...b };
	const keys = Object.keys(left);
	return keys.length === Object.keys(right).length && keys.every((key) => left[key] === right[key]);
}
