import { createBuildingBlockRequests } from 'ooxml-ui/pptx';
import { flushSync } from 'svelte';
import { expect, test, vi } from 'vitest';

import { controllerState } from '../viewer/state/controller-state';
import { routeDialogFlag } from './route-dialog-flag.svelte';

/**
 * A dialog flag raised as a host request through a real session request store, read through the
 * binding's store adapter: the flag is handled once, even when an effect reading the store reacts
 * to the request (the loop React 18 hit with flag effects keyed on more than the flag).
 */
test('a dialog flag raises one request through the session store', () => {
	const onRequest = vi.fn();
	const requests = createBuildingBlockRequests({
		state: () => ({ available: true, needsHostUi: true }),
		onRequest,
	});
	const pending = controllerState(requests);
	let open = $state(false);
	let seen = 0;
	const stop = $effect.root(() => {
		// Something that reacts to the store, as a template or effect would.
		$effect(() => {
			void pending.current;
			seen += 1;
		});
		routeDialogFlag(
			() => open,
			() => {
				open = false;
				requests.request({ kind: 'smartArt' });
			},
		);
	});
	flushSync();
	open = true;
	flushSync();
	expect(onRequest).toHaveBeenCalledOnce();
	expect(open).toBe(false);
	expect(requests.getSnapshot()).toEqual({ kind: 'smartArt' });
	expect(seen).toBeGreaterThan(1);
	open = true;
	flushSync();
	expect(onRequest).toHaveBeenCalledTimes(2);
	stop();
});
