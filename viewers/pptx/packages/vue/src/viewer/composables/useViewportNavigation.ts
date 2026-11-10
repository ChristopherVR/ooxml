/**
 * useViewportNavigation: Ctrl+wheel and pinch zoom at the pointer, and drag to
 * pan, in the slide viewport.
 *
 * The behaviour is the shared `attachViewportNavigation` (one definition for
 * all five bindings); this composable only follows the scrolling element and
 * detaches with the component. The slide inside it is looked up on each use,
 * so a remounted canvas needs no re-attach. Showing another slide keeps the
 * zoom and scrolls the viewport to the new slide's top-left corner (the shared
 * `resetViewportScroll` policy).
 */
import { attachViewportNavigation, resetViewportScroll } from 'ooxml-ui/pptx';
import type { ViewerZoomStore } from 'ooxml-ui/pptx';
import type { Ref } from 'vue';
import { onBeforeUnmount, watch } from 'vue';

export interface UseViewportNavigationOptions {
	/** The element that scrolls the slide (and contains it), or null while unmounted. */
	viewport: Ref<HTMLElement | null>;
	/** The viewer's zoom store: read for the current factor, written for a zoom. */
	store: ViewerZoomStore;
	/** True when a plain drag has nothing to select and should pan (read-only view). */
	dragPans: Ref<boolean>;
	/** False while the viewport must leave the wheel and pointer alone (a running show). */
	active: Ref<boolean>;
	/** Id of the slide on the canvas; another slide is shown from its top-left corner. */
	slideId: Ref<string | undefined>;
}

export function useViewportNavigation(options: UseViewportNavigationOptions): void {
	let detach: (() => void) | null = null;
	watch(
		options.viewport,
		(viewport) => {
			detach?.();
			detach = viewport
				? attachViewportNavigation(viewport, {
						getZoom: () => options.store.getState().zoom,
						setZoom: (zoom) => options.store.dispatch({ type: 'set-zoom', zoom }),
						dragPans: () => options.dragPans.value,
						isActive: () => options.active.value,
					})
				: null;
		},
		{ immediate: true, flush: 'post' },
	);
	watch(
		options.slideId,
		() => {
			if (options.viewport.value) {
				resetViewportScroll(options.viewport.value);
			}
		},
		{ flush: 'post' },
	);
	onBeforeUnmount(() => {
		detach?.();
		detach = null;
	});
}
