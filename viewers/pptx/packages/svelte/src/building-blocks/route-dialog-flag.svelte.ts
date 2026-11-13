import { untrack } from 'svelte';

/**
 * Clear a dialog flag and raise its host request when it turns on. The effect tracks the flag
 * alone and runs `handle` untracked: raising notifies the session's store, and an effect that also
 * subscribed to what that notification changes could run again while the flag is still on (the
 * loop React 18 hit). `handle` clears the flag before it raises. Call it during component or
 * `$effect.root` setup.
 */
export function routeDialogFlag(read: () => boolean, handle: () => void): void {
	$effect(() => {
		if (read()) {
			untrack(handle);
		}
	});
}
