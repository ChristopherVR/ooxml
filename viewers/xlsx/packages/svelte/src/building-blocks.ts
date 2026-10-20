import { onDestroy } from 'svelte';
import { writable } from 'svelte/store';
import {
	createBuildingBlockSession,
	emptyBuildingBlockState,
	type BuildingBlockOptions,
} from '../../bindings/src/building-blocks';
export type {
	BuildingBlockOptions,
	BuildingBlockState,
	SpreadsheetToolbarProps,
} from '../../bindings/src/building-blocks';
export { toolbarButtons, selectReference } from '../../bindings/src/building-blocks';
/** Create during component initialization; use $blocks for state and $effect(() => blocks.update(options)). */
export function createViewerBuildingBlocks(initial: BuildingBlockOptions = {}) {
	const store = writable(emptyBuildingBlockState());
	const session = createBuildingBlockSession(store.set);
	session.update(initial);
	onDestroy(session.destroy);
	return {
		subscribe: store.subscribe,
		update: session.update,
		handle: session.handle,
		canvasProps: { mount: session.mount },
		destroy: session.destroy,
	};
}
