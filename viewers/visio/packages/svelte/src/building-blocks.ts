import { onDestroy } from 'svelte';
import { derived, get, writable, type Readable } from 'svelte/store';
import {
	createBuildingBlockSession,
	effectiveMode,
	type BuildingBlockOptions,
} from '../../bindings/src/building-blocks';
/** Call during component initialization. Reactive option changes use blocks.update(options). */
export function createViewerBuildingBlocks(
	options: BuildingBlockOptions | Readable<BuildingBlockOptions> = {},
) {
	const source = 'subscribe' in options ? options : undefined;
	const initial = source ? get(source) : (options as BuildingBlockOptions);
	// Permission is applied at creation, before the canvas can mount or be used.
	const session = createBuildingBlockSession(initial);
	const state = writable(session.controller.state);
	const latest = writable<BuildingBlockOptions>(initial);
	const stop = session.controller.subscribe((value) => state.set(value));
	let stopOptions = () => {};
	onDestroy(() => {
		stopOptions();
		stop();
		session.destroy();
	});
	const update = (next: BuildingBlockOptions) => {
		latest.set(next);
		session.update(next);
	};
	if (source) stopOptions = source.subscribe(update);
	else update(initial);
	return {
		controller: session.controller,
		state: { subscribe: state.subscribe },
		toolbarProps: derived([state, latest], ([, next]) => session.toolbar(next)),
		loading: derived(state, (value) => value.loading),
		error: derived(state, (value) => value.error ?? value.edit.error),
		/** Effective permission: host `canEdit` and a drawing that can be saved back. */
		mode: derived(state, (value) => effectiveMode(value)),
		canvasProps: { mount: session.mount },
		inspectorProps: { controller: session.controller },
		shapesProps: derived(latest, (next) => ({
			controller: session.controller,
			locale: next.locale,
		})),
		/** Acknowledge a host save: clears dirty and fires `onDirtyChange(false)`. */
		markClean: () => session.markClean(),
		update,
	};
}
export type {
	BuildingBlockOptions,
	BuildingBlockHostCallbacks,
	DiagramToolbarProps,
} from '../../bindings/src/building-blocks';
export type { ToolbarGroupId } from '../../bindings/src/toolbar-controls';
