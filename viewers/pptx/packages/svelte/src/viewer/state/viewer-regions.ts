import { ViewerRegionHost } from 'ooxml-ui/pptx';
import type { ViewerRegionContext, ViewerRegionRenderer } from 'ooxml-ui/pptx';
import { getContext, setContext } from 'svelte';

/** The part of a region context the viewer root supplies; a region site adds `region` and `close`. */
export type ViewerRegionBase = Omit<ViewerRegionContext, 'region' | 'close'>;

const REGION_BASE_KEY = Symbol('pptx-svelte-region-base');

/** Root wiring: publish the getter that builds the base region context from viewer state. */
export function provideRegionBase(getBase: () => ViewerRegionBase): void {
	setContext(REGION_BASE_KEY, getBase);
}

/** The base context getter, or undefined when mounted stand-alone (no root). */
export function useRegionBase(): (() => ViewerRegionBase) | undefined {
	return getContext<(() => ViewerRegionBase) | undefined>(REGION_BASE_KEY);
}

export interface ViewerRegionActionParams {
	renderer: ViewerRegionRenderer;
	context: ViewerRegionContext;
}

/**
 * Svelte action hosting one region: creates a `ViewerRegionHost` on the node,
 * renders on mount and whenever the params change, destroys on unmount.
 */
export function viewerRegionHost(node: HTMLElement, params: ViewerRegionActionParams) {
	const host = new ViewerRegionHost(node);
	host.render(params.renderer, params.context);
	return {
		update(next: ViewerRegionActionParams): void {
			host.render(next.renderer, next.context);
		},
		destroy(): void {
			host.destroy();
		},
	};
}
