import { effect, untracked } from '@angular/core';
import type { EffectRef, Injector } from '@angular/core';

/**
 * Clear a dialog flag and raise its host request when it turns on. The effect tracks the flag
 * alone and runs `handle` untracked: raising notifies the session's store, and an effect that also
 * read what that notification changes could run again while the flag is still on (the loop React
 * 18 hit). `handle` clears the flag before it raises. Call it in an injection context or pass one.
 */
export function routeDialogFlag(
	flag: () => boolean,
	handle: () => void,
	injector?: Injector,
): EffectRef {
	return effect(
		() => {
			if (flag()) {
				untracked(handle);
			}
		},
		injector ? { injector } : undefined,
	);
}
