/**
 * Builds the region context (`ViewerRegionContext`) for every host-replaced
 * region from the live viewer state. One place, so the inspector, the dialogs
 * and the File tab pages all hand a host the same fields (the shared Options
 * dialog binding adds the `dialog.options` fields itself).
 *
 * @module viewer/region-support
 */
import type {
	HostBackstageRegionId,
	PowerPointViewerAPI,
	ResolvedCustomization,
	ViewerOptionsStore,
	ViewerRegionContext,
	ViewerRegionId,
} from 'ooxml-ui/pptx';

import type { Store, ViewerState } from './state';
import { isReadOnlyLocked } from './state/read-only-session';
import type { RegionSupport } from './ui/region-set';

export interface RegionSupportDeps {
	/** The object that implements `PowerPointViewerAPI` on the public instance. */
	viewer: PowerPointViewerAPI;
	store: Store<ViewerState>;
	/** The File > Options store; undefined while the viewer is still being constructed. */
	getOptionsStore(): ViewerOptionsStore | undefined;
	getResolved(): ResolvedCustomization;
	getLocale(): string;
}

export function createRegionSupport(deps: RegionSupportDeps): RegionSupport {
	const baseContext = (
		region: ViewerRegionId | HostBackstageRegionId,
		close: () => void,
	): ViewerRegionContext => {
		const state = deps.store.get();
		return {
			region,
			viewer: deps.viewer,
			mode: deps.viewer.getMode(),
			canEdit: state.editable && !state.protectedView && !isReadOnlyLocked(state),
			activeSlideIndex: state.currentSlide,
			slideCount: state.slides.length,
			selectedElementIds: state.selectedElementIds,
			locale: deps.getLocale(),
			close,
		};
	};

	return {
		getResolved: () => deps.getResolved(),
		buildContext: baseContext,
		subscribe(listener) {
			// The chrome (and its panels) is built before the options store exists,
			// so the options subscription is made as soon as the store is there.
			let offOptions: (() => void) | undefined;
			const attachOptions = (): void => {
				offOptions ??= deps.getOptionsStore()?.subscribe(listener);
			};
			attachOptions();
			const offState = deps.store.subscribe(() => {
				attachOptions();
				listener();
			});
			return () => {
				offState();
				offOptions?.();
			};
		},
	};
}
