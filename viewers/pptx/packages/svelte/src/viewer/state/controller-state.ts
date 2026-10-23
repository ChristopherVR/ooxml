import type { PptxController } from 'ooxml-ui/pptx';
import { createSubscriber } from 'svelte/reactivity';

/** A controller snapshot as a reactive value: read `current` in a template or rune. */
export interface ControllerState<S> {
	readonly current: S;
}

/**
 * The Svelte store adapter for the shared controller convention (`getSnapshot`
 * / `subscribe` / typed actions, see `render/controller` in `ooxml-ui/pptx`).
 * One adapter serves every controller: reading `current` inside an effect or a
 * template subscribes to the controller for as long as that reader lives, and a
 * read anywhere else simply returns the latest snapshot.
 */
export function controllerState<S>(controller: PptxController<S>): ControllerState<S> {
	const track = createSubscriber((update) => controller.subscribe(update));
	return {
		get current(): S {
			track();
			return controller.getSnapshot();
		},
	};
}
