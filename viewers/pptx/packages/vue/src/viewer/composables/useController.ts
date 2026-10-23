import type { PptxController } from 'ooxml-ui/pptx';
import { getCurrentScope, onScopeDispose, shallowRef } from 'vue';
import type { ShallowRef } from 'vue';

/**
 * The Vue store adapter for the shared controller convention (`getSnapshot` /
 * `subscribe` / typed actions, see `render/controller` in `ooxml-ui/pptx`).
 * One adapter serves every controller: the snapshot becomes a shallow ref that
 * follows every controller change, and the subscription ends with the current
 * effect scope (a component's setup, or an `effectScope`).
 */
export function useController<S>(controller: PptxController<S>): Readonly<ShallowRef<S>> {
	const state = shallowRef(controller.getSnapshot());
	const unsubscribe = controller.subscribe(() => {
		state.value = controller.getSnapshot();
	});
	if (getCurrentScope()) {
		onScopeDispose(unsubscribe);
	}
	return state;
}
