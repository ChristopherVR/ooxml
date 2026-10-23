import { signal } from '@angular/core';
import type { DestroyRef, Signal } from '@angular/core';
import type { PptxController } from 'ooxml-ui/pptx';

/**
 * The Angular store adapter for the shared controller convention
 * (`getSnapshot` / `subscribe` / typed actions, see `render/controller` in
 * `ooxml-ui/pptx`). One adapter serves every controller: the snapshot becomes a
 * read-only signal that updates on every controller change. Pass a
 * `DestroyRef` when the controller outlives the owner; a controller created by
 * the same service needs none, since both are collected together.
 */
export function controllerSignal<S>(
	controller: PptxController<S>,
	destroyRef?: DestroyRef | null,
): Signal<S> {
	const state = signal(controller.getSnapshot());
	const unsubscribe = controller.subscribe(() => state.set(controller.getSnapshot()));
	destroyRef?.onDestroy(unsubscribe);
	return state.asReadonly();
}
