import { effect, untracked } from '@angular/core';

/**
 * Calls `report(value)` each time `read()` changes after the component was created, never for
 * the value it was created with. Host outputs that describe a transition (autosave status,
 * slide show, connection status) use it so a host hears about changes, not about the initial
 * state.
 *
 * Call it in an injection context (a constructor or field initialiser).
 */
export function reportOnChange<T>(read: () => T, report: (value: T) => void): void {
	let last: { value: T } | undefined;
	effect(() => {
		const value = read();
		if (last === undefined) {
			last = { value };
			return;
		}
		if (Object.is(last.value, value)) {
			return;
		}
		last.value = value;
		untracked(() => report(value));
	});
}
