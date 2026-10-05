/**
 * Index access helper for `noUncheckedIndexedAccess`. Use only where the index is guaranteed
 * to be in range by the surrounding algorithm; throws if the invariant is ever violated.
 */
export function at<T>(items: ArrayLike<T>, index: number): T {
	const value = items[index];
	if (value === undefined) {
		throw new RangeError(`Index ${index} out of range (length ${items.length})`);
	}
	return value;
}
