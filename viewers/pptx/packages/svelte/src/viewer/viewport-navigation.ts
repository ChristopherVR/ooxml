import { attachViewportNavigation } from 'ooxml-ui/pptx';

export interface ViewportNavigationParams {
	/** Current manual zoom factor (1 = fit). */
	getScale: () => number;
	/** Commit a new manual zoom factor. */
	setScale: (scale: number) => void;
	/** True when a plain drag has nothing to select and should pan (read-only view). */
	dragPans: boolean;
	/** False while the viewport must leave the wheel and pointer alone (a running show). */
	active: boolean;
}

/**
 * Svelte action: Ctrl+wheel and pinch zoom at the pointer, and drag to pan, in
 * the slide viewport. The behaviour is the shared `attachViewportNavigation`
 * (one definition for all five bindings); parameters are read live, so an
 * update needs no re-attach.
 */
export function viewportNavigation(node: HTMLElement, initial: ViewportNavigationParams) {
	let params = initial;
	const detach = attachViewportNavigation(node, {
		getZoom: () => params.getScale(),
		setZoom: (zoom) => params.setScale(zoom),
		dragPans: () => params.dragPans,
		isActive: () => params.active,
	});
	return {
		update(next: ViewportNavigationParams) {
			params = next;
		},
		destroy: detach,
	};
}
