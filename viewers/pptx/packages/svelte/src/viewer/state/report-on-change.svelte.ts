import { untrack } from 'svelte';

/**
 * Calls `report(value)` each time `read()` changes after the component mounted, never for the
 * value it mounted with. Host events that describe a transition (autosave status, slide show,
 * connection status) use it so a host hears about changes, not about the initial state.
 *
 * Call it during component initialisation, like any `$effect`.
 */
export function reportOnChange<T>(read: () => T, report: (value: T) => void): void {
	let last: { value: T } | undefined;
	$effect(() => {
		const value = read();
		if (last === undefined) {
			last = { value };
			return;
		}
		if (Object.is(last.value, value)) {
			return;
		}
		last.value = value;
		untrack(() => report(value));
	});
}
