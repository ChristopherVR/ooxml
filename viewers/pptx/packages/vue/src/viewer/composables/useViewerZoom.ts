/**
 * useViewerZoom: the editor canvas's two independent scale factors.
 *
 * `zoom` is the user's zoom level (the status bar's +/- and Ctrl+scroll).
 * `fitScale` is the fit-to-viewport factor (<= 1) that `SlideCanvas` measures
 * with a ResizeObserver so a whole slide is visible by default instead of
 * overflowing a small viewport. Everything that renders or maps pointer
 * coordinates must use `effectiveZoom` (their product), matching React and
 * Angular where "100%" means "fit to viewport".
 *
 * The state lives in the shared `createViewerZoomStore`, so the zoom MODEL (not
 * just the step size) is one definition across all five bindings; this
 * composable is the Vue projection of it. Both factors stay writable refs so
 * existing call sites (`fitScale = $event` from `SlideCanvas`, the pinch
 * gesture, the public `setZoom` API) are unchanged; the writes are simply
 * routed through the store's semantic commands, which also means a pinch is now
 * clamped by the same bounds as every other zoom entry point.
 */
import { createViewerZoomStore } from 'ooxml-ui/pptx';
import type { ViewerZoomStore } from 'ooxml-ui/pptx';
import type { ComputedRef, Ref, WritableComputedRef } from 'vue';
import { computed, watch } from 'vue';

import { useViewerStore } from './useViewerStore';

export interface UseViewerZoomResult {
	/** The user's zoom level (1 = fit to viewport). */
	zoom: Ref<number>;
	/** Fit-to-viewport factor reported by `SlideCanvas`'s ResizeObserver. */
	fitScale: Ref<number>;
	/** Effective on-screen scale: `fitScale x zoom`. */
	effectiveZoom: ComputedRef<number>;
	zoomIn: () => void;
	zoomOut: () => void;
	zoomReset: () => void;
	/**
	 * Report the slide about to be shown, before the watcher sees it, so a zoom
	 * set right after (the handle's `goTo` then `setZoom`) is not undone.
	 */
	showSlide: (slideId: string | undefined) => void;
	/** The underlying store, for callers that want a narrower subscription. */
	store: ViewerZoomStore;
}

/**
 * @param activeSlideId Id of the slide on the canvas. Going to another slide
 *   returns the zoom to fit; the shared store owns that policy, and this only
 *   reports the slide shown.
 */
export function useViewerZoom(activeSlideId?: () => string | undefined): UseViewerZoomResult {
	const store = createViewerZoomStore();
	const showSlide = (slideId: string | undefined): void =>
		store.dispatch({ type: 'slide-shown', slideId });
	if (activeSlideId) {
		watch(activeSlideId, showSlide, { immediate: true });
	}
	// Two separate selector subscriptions rather than one on the whole state: a
	// viewport re-measure then wakes only what reads `fitScale`.
	const zoomValue = useViewerStore(store, (state) => state.zoom);
	const fitValue = useViewerStore(store, (state) => state.fitScale);

	const zoom: WritableComputedRef<number> = computed({
		get: () => zoomValue.value,
		set: (value) => store.dispatch({ type: 'set-zoom', zoom: value }),
	});
	const fitScale: WritableComputedRef<number> = computed({
		get: () => fitValue.value,
		set: (value) => store.dispatch({ type: 'set-fit-scale', fitScale: value }),
	});

	return {
		zoom,
		fitScale,
		effectiveZoom: computed(() => fitValue.value * zoomValue.value),
		zoomIn: () => store.dispatch({ type: 'zoom-in' }),
		zoomOut: () => store.dispatch({ type: 'zoom-out' }),
		zoomReset: () => store.dispatch({ type: 'zoom-to-fit' }),
		showSlide,
		store,
	};
}
