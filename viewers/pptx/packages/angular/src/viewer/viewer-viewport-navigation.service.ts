/**
 * viewer-viewport-navigation.service.ts: Ctrl+wheel and pinch zoom at the
 * pointer, and drag to pan, in the slide viewport.
 *
 * The behaviour is the shared `attachViewportNavigation` (one definition for
 * all five bindings); this service only attaches it to `<main>`, the element
 * that scrolls the slide in this binding, and detaches it with the component.
 * `<main>` renders only once a deck has loaded, so the service follows the
 * element query rather than looking once after the first render. The accessors
 * are read live, so one attach handles every mode.
 *
 * Provide it once on the viewer component (`providers: [ViewerViewportNavigationService]`).
 */

import { DestroyRef, effect, inject, Injectable, Injector, untracked } from '@angular/core';
import { attachViewportNavigation } from 'ooxml-ui/pptx';

import { ViewerZoomService } from './viewer-zoom.service';

/** Live host accessors the controller consults. */
interface ViewportNavigationHost {
	/** True when a plain drag has nothing to select and should pan (read-only view). */
	readonly dragPans: () => boolean;
	/** True while a show is running and owns the wheel and pointer. */
	readonly presenting: () => boolean;
}

@Injectable()
export class ViewerViewportNavigationService {
	private readonly zoomSvc = inject(ViewerZoomService);
	private readonly destroyRef = inject(DestroyRef);
	private readonly injector = inject(Injector);

	/**
	 * Attach to `mainEl()`'s element whenever it is in the DOM, moving to a
	 * replacement element (called once from the host component's constructor).
	 * `mainEl` must read the element query signal, so a change re-runs it.
	 */
	setup(mainEl: () => HTMLElement | undefined, host: ViewportNavigationHost): void {
		let attached: { el: HTMLElement; detach: () => void } | null = null;
		const sync = (el: HTMLElement | null): void => {
			if (attached?.el === el) {
				return;
			}
			attached?.detach();
			attached = el
				? {
						el,
						detach: attachViewportNavigation(el, {
							getZoom: () => this.zoomSvc.zoom(),
							setZoom: (zoom) => this.zoomSvc.setZoom(zoom),
							dragPans: () => host.dragPans(),
							isActive: () => !host.presenting(),
						}),
					}
				: null;
		};
		effect(
			() => {
				const el = mainEl() ?? null;
				untracked(() => sync(el));
			},
			{ injector: this.injector },
		);
		this.destroyRef.onDestroy(() => sync(null));
	}
}
