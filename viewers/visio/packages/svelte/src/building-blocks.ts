import { onDestroy } from 'svelte';
import { derived, get, writable, type Readable } from 'svelte/store';
import {
	createBuildingBlockSession,
	type BuildingBlockOptions,
} from '../../bindings/src/building-blocks';
/** Call during component initialization. Reactive option changes use blocks.update(options). */
export function createViewerBuildingBlocks(
	options: BuildingBlockOptions | Readable<BuildingBlockOptions> = {},
) {
	const source = 'subscribe' in options ? options : undefined;
	const initial = source ? get(source) : (options as BuildingBlockOptions);
	const session = createBuildingBlockSession();
	const state = writable(session.controller.state);
	const locale = writable<string | undefined>(initial.locale);
	const stop = session.controller.subscribe((value) => state.set(value));
	let stopOptions = () => {};
	onDestroy(() => {
		stopOptions();
		stop();
		session.destroy();
	});
	const update = (next: BuildingBlockOptions) => {
		locale.set(next.locale);
		session.update(next);
	};
	if (source) stopOptions = source.subscribe(update);
	else update(initial);
	return {
		controller: session.controller,
		state: { subscribe: state.subscribe },
		toolbarProps: derived([state, locale], ([, language]) => session.toolbar(language)),
		loading: derived(state, (value) => value.loading),
		error: derived(state, (value) => value.error ?? value.edit.error),
		mode: derived(state, (value) => (value.canEdit ? ('edit' as const) : ('view' as const))),
		canvasProps: { mount: session.mount },
		inspectorProps: { controller: session.controller },
		update,
	};
}
export type { BuildingBlockOptions, DiagramToolbarProps } from '../../bindings/src/building-blocks';
