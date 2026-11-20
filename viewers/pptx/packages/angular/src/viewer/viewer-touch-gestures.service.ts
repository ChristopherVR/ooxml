/**
 * viewer-touch-gestures.service.ts: Wires the framework-agnostic touch-gesture
 * recogniser ({@link attachTouchGestures}) to the main canvas host. Mirrors
 * React's `useTouchGestures` wiring:
 *   - pinch-to-zoom always updates {@link ViewerZoomService}'s zoom signal
 *     (clamped to the viewer range), with `preventDefault()` on the pinch path
 *     to suppress the browser's native pinch-zoom;
 *   - horizontal swipe navigates slides, but only when editing is off
 *     (`!canEdit()`): in edit mode single-finger gestures belong to element
 *     manipulation (move/resize/rotate), so we never hijack them. The large
 *     ‹ › buttons remain available for explicit navigation in all modes;
 *   - long-press and double-tap are not handled here: {@link setupTouchPress}
 *     installs the shared `installTouchPress` on the viewer root, which turns
 *     them into `contextmenu` / `dblclick` on the controls that opt in with
 *     `data-pptx-touch-press` (issue #290), the canvas included.
 *
 * Extracted from {@link PowerPointViewerComponent}: the component calls
 * {@link setup} once from its constructor with the `<main>` element accessor and
 * the mode/navigation accessors the callbacks gate on. The recogniser's swipe/
 * long-press callbacks check the live accessors, so a single attach handles
 * every mode without re-binding.
 *
 * Provide it once on the viewer component (`providers: [ViewerTouchGesturesService]`).
 */

import { afterNextRender, DestroyRef, inject, Injectable } from '@angular/core';
import { installTouchPress } from 'ooxml-ui/pptx';

import { attachTouchGestures } from './touch-gestures';
import { ViewerZoomService } from './viewer-zoom.service';

/** Live host accessors the gesture recogniser consults. */
interface TouchGesturesHost {
	readonly canEdit: () => boolean;
	readonly presenting: () => boolean;
	readonly goPrev: () => void;
	readonly goNext: () => void;
}

@Injectable()
export class ViewerTouchGesturesService {
	private readonly zoomSvc = inject(ViewerZoomService);
	private readonly destroyRef = inject(DestroyRef);

	/**
	 * Attach the recogniser to `mainEl()`'s element once it is in the DOM
	 * (called once from the host component's constructor).
	 */
	setup(mainEl: () => HTMLElement | undefined, host: TouchGesturesHost): void {
		afterNextRender(() => {
			const el = mainEl();
			if (!el) {
				return;
			}
			const teardown = attachTouchGestures(el, {
				getScale: () => this.zoomSvc.zoom(),
				callbacks: {
					onPinchZoom: (newScale) => this.zoomSvc.setZoom(newScale),
					onSwipe: (direction) => {
						// Edit mode: leave single-finger gestures to element manipulation.
						if (host.canEdit()) {
							return;
						}
						// direction 1 = swipe right (previous), -1 = swipe left (next).
						if (direction === 1) {
							host.goPrev();
						} else {
							host.goNext();
						}
					},
				},
			});
			this.destroyRef.onDestroy(teardown);
		});
	}

	/**
	 * Turn a touch long-press into `contextmenu` and a double-tap into `dblclick`
	 * under `rootEl()` (iOS Safari fires neither; issue #290).
	 */
	setupTouchPress(rootEl: () => HTMLElement | undefined): void {
		afterNextRender(() => {
			const el = rootEl();
			if (el) {
				this.destroyRef.onDestroy(installTouchPress(el));
			}
		});
	}
}
