/**
 * useViewerRegions: the host-supplied UI regions (the "replace" half of the
 * customisation model) for the Vue binding.
 *
 * `PowerPointViewer` provides the base context (everything a renderer needs
 * except which region it fills and how to dismiss it) under
 * {@link ViewerRegionBaseKey}. A site that can be replaced calls
 * {@link useRegion}: it gets the host renderer for its region id (or
 * `undefined` to render the built-in) and a builder for the full context.
 * The mount, update and teardown semantics live in the shared
 * `ViewerRegionHost`, driven by `ViewerRegion.vue`.
 */
import { getRegionRenderer } from 'ooxml-ui/pptx';
import type {
	HostBackstageRegionId,
	ViewerRegionContext,
	ViewerRegionId,
	ViewerRegionRenderer,
} from 'ooxml-ui/pptx';
import { computed, inject } from 'vue';
import type { ComputedRef, InjectionKey } from 'vue';

import { useResolvedCustomization } from './useViewerCustomization';

/** The viewer-wide part of a region context. */
export type ViewerRegionBase = Omit<ViewerRegionContext, 'region' | 'close'>;

/** Provided by `PowerPointViewer`; absent when a sub-component is mounted alone. */
export const ViewerRegionBaseKey: InjectionKey<ComputedRef<ViewerRegionBase>> =
	Symbol('pptxViewerRegionBase');

export interface UseRegionResult {
	/** The host renderer replacing this region; `undefined` renders the built-in. */
	renderer: ComputedRef<ViewerRegionRenderer | undefined>;
	/** The full context for this region; `extra` carries region specific fields. */
	context: (close: () => void, extra?: object) => ViewerRegionContext | undefined;
}

type RegionIdSource =
	| ViewerRegionId
	| HostBackstageRegionId
	| (() => ViewerRegionId | HostBackstageRegionId | undefined);

/** Look up the renderer for `region` (an id, or a getter for a prop-driven id). */
export function useRegion(source: RegionIdSource): UseRegionResult {
	const idOf = typeof source === 'function' ? source : () => source;
	const resolved = useResolvedCustomization();
	const base = inject(ViewerRegionBaseKey, undefined);
	const renderer = computed(() => {
		const region = idOf();
		return base && region && !region.startsWith('backstage.host.')
			? getRegionRenderer(resolved.value, region as ViewerRegionId)
			: undefined;
	});
	return {
		renderer,
		context: (close, extra) => {
			const region = idOf();
			return base && region
				? ({ ...base.value, ...extra, region, close } as ViewerRegionContext)
				: undefined;
		},
	};
}

/** The base context for a host File tab page (not a built-in region id). */
export function useRegionBase(): ComputedRef<ViewerRegionBase> | undefined {
	return inject(ViewerRegionBaseKey, undefined);
}
