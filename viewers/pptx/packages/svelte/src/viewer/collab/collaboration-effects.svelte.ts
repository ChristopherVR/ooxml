/**
 * The `$effect`s the collaboration controller registers on construction,
 * lifted out of `collaboration.svelte.ts` to keep that file inside the repo's
 * 300 LOC ceiling. They are pure wiring: every piece of state they touch is
 * reached through the {@link CollaborationEffectsHost} callbacks, so this module
 * holds no state of its own.
 */

import type { CollaborationConfig } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';

export interface CollaborationEffectsHost {
	getConfig: () => CollaborationConfig | undefined;
	getSlides: () => PptxSlide[];
	syncConfig: (config: CollaborationConfig | undefined) => void;
	isPublishable: () => boolean;
	flushLocalSlides: (slides: PptxSlide[]) => void;
	/** Route Undo and Redo through the room's history while it is writable (returns the release). */
	bindHistory?: () => (() => void) | void;
	/** Tear the session down with the component (the shared session controller also leaves on `pagehide`). */
	destroy: () => void;
}

/**
 * Register the controller's effects. Call from the constructor, inside the
 * component's reactive scope.
 */
export function registerCollaborationEffects(host: CollaborationEffectsHost): void {
	// Auto start/stop when the host supplies (or clears) a config. Compared by
	// reference so re-emitting the same object does not restart the session.
	$effect(() => {
		host.syncConfig(host.getConfig());
	});

	// Broadcast local slide edits granularly. Reading the active flag re-runs the
	// effect on (de)activation; the gate + role checks live in the flush.
	$effect(() => {
		const slides = host.getSlides();
		if (host.isPublishable()) {
			host.flushLocalSlides(slides);
		}
	});

	$effect(() => host.bindHistory?.());

	// Leave the room and drop the pagehide listeners when the component goes.
	$effect(() => () => host.destroy());
}
