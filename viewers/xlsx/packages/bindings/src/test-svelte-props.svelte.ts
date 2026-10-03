// Test support (imported by svelte.test.ts only): a `$state` proxy, so a test can change the
// props of a mounted component the way a parent re-render would.
export function reactiveProps<T extends object>(initial: T): T {
	const props = $state(initial);
	return props;
}
