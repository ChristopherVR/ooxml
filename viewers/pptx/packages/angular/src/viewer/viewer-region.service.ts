/**
 * viewer-region.service.ts: builds the {@link ViewerRegionContext} every
 * replaced region receives, and looks up the renderer a host registered.
 *
 * The root component attaches a getter for the viewer state once; render sites
 * (the root template, the dialogs) ask for `renderer(id)` and `context(...)`.
 * Provided per viewer through `POWER_POINT_VIEWER_PROVIDERS`; every consumer
 * injects it optionally so components keep working outside a viewer.
 */
import { inject, Injectable } from '@angular/core';
import { EMPTY_RESOLVED_CUSTOMIZATION, getRegionRenderer } from 'ooxml-ui/pptx';
import type {
	HostBackstageRegionId,
	PowerPointViewerAPI,
	ViewerMode,
	ViewerOptionPrimitive,
	ViewerRegionContext,
	ViewerRegionId,
	ViewerRegionRenderer,
	ViewerSettingsRegionContext,
	OptionsSettingId,
	ViewerOptionsGroupId,
} from 'ooxml-ui/pptx';

import { ViewerCustomizationService } from './viewer-customization.service';
import { ViewerOptionsService } from './viewer-options.service';

/** The viewer state a region context is built from (read inside a reactive context). */
export interface ViewerRegionBase {
	viewer: PowerPointViewerAPI;
	mode: ViewerMode;
	canEdit: boolean;
	activeSlideIndex: number;
	slideCount: number;
	selectedElementIds: readonly string[];
	locale: string;
}

@Injectable()
export class ViewerRegionService {
	private readonly customization = inject(ViewerCustomizationService, { optional: true });
	private readonly options = inject(ViewerOptionsService, { optional: true });
	private source: (() => ViewerRegionBase) | undefined;

	/** Called once by the root component with a getter over its signals. */
	attach(source: () => ViewerRegionBase): void {
		this.source = source;
	}

	/**
	 * The host renderer replacing `id`, or undefined for the built-in. Also
	 * undefined while no context can be built, so a site never renders an empty body.
	 */
	renderer(id: ViewerRegionId): ViewerRegionRenderer | undefined {
		if (!this.source) {
			return undefined;
		}
		const resolved = this.customization?.resolved() ?? EMPTY_RESOLVED_CUSTOMIZATION;
		return getRegionRenderer(resolved, id);
	}

	/** The context for `region`; null until the root component attached its state. */
	context(
		region: ViewerRegionId | HostBackstageRegionId,
		close: () => void,
	): ViewerRegionContext | null {
		const base = this.source?.();
		return base ? { region, ...base, close } : null;
	}

	/** The `dialog.options` context: the base plus the options model and its setters. */
	settingsContext(close: () => void): ViewerSettingsRegionContext | null {
		const base = this.source?.();
		const store = this.options;
		if (!base || !store) {
			return null;
		}
		return {
			...base,
			region: 'dialog.options',
			close,
			options: store.options(),
			setOption: (setting: OptionsSettingId, value: ViewerOptionPrimitive) => {
				const dot = setting.indexOf('.');
				store.setValue(
					setting.slice(0, dot) as ViewerOptionsGroupId,
					setting.slice(dot + 1),
					value,
				);
			},
			resetOptions: () => store.reset(undefined),
		};
	}
}
