/**
 * Shared mounting helpers for the region customization tests. The tests are
 * split over several files, each in its own worker, because a full viewer in
 * happy-dom is heavy and one file mounting all of them exhausts the heap.
 */
import type { ViewerRegionContext } from 'ooxml-ui/pptx';

import { createPptxViewer } from './PptxViewer';
import type { Store, ViewerState } from './state';
import type { PptxViewerInstance, PptxViewerOptions } from './types';

export type { ViewerRegionContext };

let active: PptxViewerInstance[] = [];

export function mount(options: PptxViewerOptions = {}): {
	container: HTMLElement;
	viewer: PptxViewerInstance;
} {
	const container = document.createElement('div');
	document.body.appendChild(container);
	const viewer = createPptxViewer(container, options);
	active.push(viewer);
	return { container, viewer };
}

/** Destroy every viewer `mount` created; register with `afterEach`. */
export function destroyMounted(): void {
	for (const viewer of active) {
		viewer.destroy();
	}
	active = [];
	document.body.replaceChildren();
}

/** The viewer state store (private on the class; used to drive selection). */
export function storeOf(viewer: PptxViewerInstance): Store<ViewerState> {
	return (viewer as unknown as { store: Store<ViewerState> }).store;
}

/** The dialog openers the class has but the public instance type leaves out. */
export function internals(viewer: PptxViewerInstance): {
	openSettings(): void;
	openShare(): void;
	openPrintDialog(): void;
	openSelectionPane(): void;
	openAccessibility(): void;
	openPasswordProtection(): void;
	openDigitalSignatures(): void;
} {
	return viewer as unknown as ReturnType<typeof internals>;
}

export function region(container: ParentNode, id: string): HTMLElement | null {
	return container.querySelector<HTMLElement>(`[data-pptx-region="${id}"]`);
}
