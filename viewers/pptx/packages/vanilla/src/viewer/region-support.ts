/**
 * Builds the region context (`ViewerRegionContext`) for every host-replaced
 * region from the live viewer state. One place, so the inspector, the dialogs
 * and the File tab pages all hand a host the same fields.
 *
 * @module viewer/region-support
 */
import type {
	HostBackstageRegionId,
	OptionsSettingId,
	PowerPointViewerAPI,
	ResolvedCustomization,
	ViewerOptionPrimitive,
	ViewerOptionsGroupId,
	ViewerOptionsStore,
	ViewerRegionContext,
	ViewerRegionId,
	ViewerSettingsRegionContext,
} from 'ooxml-ui/pptx';

import type { Store, ViewerState } from './state';
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
			canEdit:
				state.editable &&
				!state.protectedView &&
				!(state.readOnlyRecommendation?.defaultReadOnly && !state.readOnlyBannerDismissed),
			activeSlideIndex: state.currentSlide,
			slideCount: state.slides.length,
			selectedElementIds: state.selectedElementIds,
			locale: deps.getLocale(),
			close,
		};
	};

	return {
		getResolved: () => deps.getResolved(),
		buildContext(region, close) {
			const base = baseContext(region, close);
			if (region !== 'dialog.options') {
				return base;
			}
			const options = deps.getOptionsStore();
			if (!options) {
				return base;
			}
			const settings: ViewerSettingsRegionContext = {
				...base,
				region: 'dialog.options',
				options: options.getOptions(),
				// The store applies the host's locks to every write, exactly as it
				// does for the built-in dialog's controls.
				setOption: (setting: OptionsSettingId, value: ViewerOptionPrimitive) => {
					const dot = setting.indexOf('.');
					if (dot > 0) {
						options.setValue(
							setting.slice(0, dot) as ViewerOptionsGroupId,
							setting.slice(dot + 1),
							value,
						);
					}
				},
				resetOptions: () => options.reset(),
			};
			return settings;
		},
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
