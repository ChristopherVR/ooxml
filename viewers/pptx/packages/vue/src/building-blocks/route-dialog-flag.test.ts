import { createBuildingBlockRequests } from 'ooxml-ui/pptx';
import { expect, test, vi } from 'vitest';
import { effectScope, nextTick, ref, watch } from 'vue';

import { useController } from '../viewer/composables/useController';
import { routeDialogFlag } from './engine';

/**
 * A dialog flag raised as a host request through a real session request store, read through the
 * binding's store adapter: the flag is handled once, even when the request notification re-renders
 * something that watches the store (the loop React 18 hit with flag effects keyed on more than the
 * flag).
 */
test('a dialog flag raises one request through the session store', async () => {
	const onRequest = vi.fn();
	const requests = createBuildingBlockRequests({
		state: () => ({ available: true, needsHostUi: true }),
		onRequest,
	});
	const scope = effectScope();
	const flag = ref(false);
	scope.run(() => {
		const pending = useController(requests);
		// Something that reacts to the store synchronously, as a template or watcher would.
		watch(pending, () => undefined, { flush: 'sync' });
		routeDialogFlag(
			() => flag.value,
			() => {
				flag.value = false;
			},
			() => void requests.request({ kind: 'smartArt' }),
		);
	});
	flag.value = true;
	await nextTick();
	expect(onRequest).toHaveBeenCalledOnce();
	expect(flag.value).toBe(false);
	expect(requests.getSnapshot()).toEqual({ kind: 'smartArt' });
	flag.value = true;
	await nextTick();
	expect(onRequest).toHaveBeenCalledTimes(2);
	scope.stop();
});
