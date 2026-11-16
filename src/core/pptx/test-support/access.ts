// Test-only: checked accessors that keep pptx tests readable under noUncheckedIndexedAccess.
// Each helper throws a descriptive error instead of letting `undefined` leak into assertions.

/** Returns `items[index]`, failing the test with a clear message when it is missing. */
export function at<T>(items: Readonly<ArrayLike<T>> | undefined, index: number): T {
	const item = items?.[index];
	if (item === undefined) throw new Error(`Expected an item at index ${index}`);
	return item;
}

/** Narrows a defined value, failing with `what` when it is missing. */
export function must<T>(value: T | null | undefined, what = 'value'): T {
	if (value === undefined || value === null) throw new Error(`Expected ${what} to be defined`);
	return value;
}
