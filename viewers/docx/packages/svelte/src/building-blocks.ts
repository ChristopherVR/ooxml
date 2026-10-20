import { onDestroy } from 'svelte';
import { derived, readable, type Readable } from 'svelte/store';
import {
	createBuildingBlockSession,
	type BuildingBlockOptions,
	type CanvasProps,
} from '../../bindings/src/building-blocks';
export type {
	BuildingBlockOptions,
	CanvasProps,
	WordToolbarProps,
} from '../../bindings/src/building-blocks';
/** Call during component initialization. Pass a store for reactive options. */
export function createViewerBuildingBlocks(
	options: BuildingBlockOptions | Readable<BuildingBlockOptions> = {},
) {
	const session = createBuildingBlockSession();
	const stopOptions =
		'subscribe' in options ? options.subscribe((next) => session.update(next)) : undefined;
	if (!stopOptions) session.update(options as BuildingBlockOptions);
	const state = readable(session.state, (set) => session.subscribe(set));
	const destroy = () => {
		stopOptions?.();
		session.destroy();
	};
	onDestroy(destroy);
	return {
		canvasProps: session.canvasProps,
		handle: session.handle,
		toolbarProps: derived(state, (value) => value.toolbarProps),
		loading: derived(state, (value) => value.loading),
		error: derived(state, (value) => value.error),
		mode: derived(state, (value) => value.mode),
		update: (next: BuildingBlockOptions) => session.update(next),
		destroy,
	};
}
/** Svelte action for hosts supplying their own canvas markup. */
export function documentCanvas(node: HTMLElement, props: CanvasProps) {
	let current = props;
	current.mount(node);
	return {
		update(next: CanvasProps) {
			if (next.mount !== current.mount) {
				current.mount(null);
				current = next;
				current.mount(node);
			}
		},
		destroy() {
			current.mount(null);
		},
	};
}

export { prepareWordToolbar } from '../../bindings/src/building-blocks';
