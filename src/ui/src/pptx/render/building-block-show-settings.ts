import type { PptxCustomShow, PptxPresentationProperties } from 'ooxml-core/pptx';

import { BUILDING_BLOCK_OK, buildingBlockFailure } from './building-block-actions';
import type { BuildingBlockFailureCode, BuildingBlockResult } from './building-block-actions';

/**
 * Framework-free model behind Slide Show > Set Up Slide Show, for the bundled dialog and for a
 * host that renders its own (`useViewerBuildingBlocks().actions.setUpSlideShow`). One function
 * applies a change to the settings (with the dialog's rules: a kiosk show loops, a range starts
 * as the whole deck, a custom show defaults to the first), one validates the result and one
 * describes it with every default resolved. The monitor and a laser colour are not modelled by
 * the deck, so they are not offered.
 */

/** The deck facts a settings change depends on. */
export interface SlideShowSettingsContext {
	slideCount: number;
	customShows: readonly Pick<PptxCustomShow, 'id' | 'name'>[];
}

export type SlideShowType = NonNullable<PptxPresentationProperties['showType']>;
export type SlideShowAdvanceMode = NonNullable<PptxPresentationProperties['advanceMode']>;
export type SlideShowSlidesMode = NonNullable<PptxPresentationProperties['showSlidesMode']>;

/** The fields Set Up Show edits. A host passes only what it changes. */
export type SlideShowSettingsPatch = Pick<
	PptxPresentationProperties,
	| 'showType'
	| 'loopContinuously'
	| 'showWithNarration'
	| 'showWithAnimation'
	| 'showSubtitles'
	| 'advanceMode'
	| 'showSlidesMode'
	| 'showSlidesFrom'
	| 'showSlidesTo'
	| 'showSlidesCustomShowId'
	| 'penColor'
	| 'kioskRestartTime'
	| 'showScrollbar'
>;

const SHOW_TYPES: readonly SlideShowType[] = ['presented', 'browsed', 'kiosk'];
const ADVANCE_MODES: readonly SlideShowAdvanceMode[] = ['manual', 'useTimings'];
const SLIDES_MODES: readonly SlideShowSlidesMode[] = ['all', 'range', 'customShow'];

/** The settings after `patch`, with the dialog's dependent defaults. Inputs are not mutated. */
export function applySlideShowSettingsPatch(
	current: PptxPresentationProperties,
	patch: Partial<PptxPresentationProperties>,
	context: SlideShowSettingsContext,
): PptxPresentationProperties {
	const next: PptxPresentationProperties = { ...current, ...patch };
	if (patch.showType === 'kiosk') {
		next.loopContinuously = true;
	}
	if (patch.showSlidesMode === 'range') {
		next.showSlidesFrom = patch.showSlidesFrom ?? current.showSlidesFrom ?? 1;
		next.showSlidesTo = patch.showSlidesTo ?? current.showSlidesTo ?? context.slideCount;
	}
	if (patch.showSlidesMode === 'customShow') {
		const id =
			patch.showSlidesCustomShowId ?? current.showSlidesCustomShowId ?? context.customShows[0]?.id;
		if (id !== undefined) {
			next.showSlidesCustomShowId = id;
		}
	}
	return next;
}

/** Whether `settings` may be stored: known enums, a sane range and a custom show that exists. */
export function validateSlideShowSettings(
	settings: PptxPresentationProperties,
	context: SlideShowSettingsContext,
): BuildingBlockResult {
	const known = <T extends string>(value: T | undefined, list: readonly T[]) =>
		value === undefined || list.includes(value);
	if (
		!known(settings.showType, SHOW_TYPES) ||
		!known(settings.advanceMode, ADVANCE_MODES) ||
		!known(settings.showSlidesMode, SLIDES_MODES)
	) {
		return buildingBlockFailure('invalid-data');
	}
	if (settings.showSlidesMode === 'range') {
		const from = settings.showSlidesFrom ?? 1;
		const to = settings.showSlidesTo ?? context.slideCount;
		if (
			!Number.isInteger(from) ||
			!Number.isInteger(to) ||
			from < 1 ||
			to > context.slideCount ||
			from > to
		) {
			return buildingBlockFailure('invalid-range');
		}
	}
	if (
		settings.showSlidesMode === 'customShow' &&
		!context.customShows.some((show) => show.id === settings.showSlidesCustomShowId)
	) {
		return buildingBlockFailure('unknown-show');
	}
	if (settings.penColor !== undefined && !/^#?[0-9a-f]{6}$/i.test(settings.penColor)) {
		return buildingBlockFailure('invalid-data');
	}
	return BUILDING_BLOCK_OK;
}

/** The current settings with every default resolved, as a host's form needs them. */
export interface SlideShowSettingsView {
	showType: SlideShowType;
	loop: boolean;
	/** Narration is played (the model stores the negation, "show without narration"). */
	playNarration: boolean;
	playAnimation: boolean;
	subtitles: boolean;
	advanceMode: SlideShowAdvanceMode;
	slidesMode: SlideShowSlidesMode;
	rangeFrom: number;
	rangeTo: number;
	customShowId: string | undefined;
	penColor: string | undefined;
	/** Kiosk auto-restart, in milliseconds. */
	kioskRestartTime: number | undefined;
	/** Browsed shows: show the scrollbar (the schema default is true). */
	showScrollbar: boolean;
	slideCount: number;
	customShows: { id: string; name: string }[];
}

export function describeSlideShowSettings(
	settings: PptxPresentationProperties,
	context: SlideShowSettingsContext,
): SlideShowSettingsView {
	return {
		showType: settings.showType ?? 'presented',
		loop: Boolean(settings.loopContinuously),
		playNarration: settings.showWithNarration !== false,
		playAnimation: settings.showWithAnimation !== false,
		subtitles: Boolean(settings.showSubtitles),
		advanceMode: settings.advanceMode ?? 'useTimings',
		slidesMode: settings.showSlidesMode ?? 'all',
		rangeFrom: settings.showSlidesFrom ?? 1,
		rangeTo: settings.showSlidesTo ?? context.slideCount,
		customShowId: settings.showSlidesCustomShowId,
		penColor: settings.penColor,
		kioskRestartTime: settings.kioskRestartTime,
		showScrollbar: settings.showScrollbar ?? true,
		slideCount: context.slideCount,
		customShows: context.customShows.map(({ id, name }) => ({ id, name })),
	};
}

/**
 * Apply `patch` to `current`: the new settings to store, or the code they are refused with. The
 * caller stores them and marks the deck dirty.
 */
export function resolveSlideShowSettings(
	current: PptxPresentationProperties,
	patch: SlideShowSettingsPatch,
	context: SlideShowSettingsContext,
):
	| { ok: true; settings: PptxPresentationProperties }
	| { ok: false; code: BuildingBlockFailureCode } {
	const settings = applySlideShowSettingsPatch(current, patch, context);
	const checked = validateSlideShowSettings(settings, context);
	return checked.ok ? { ok: true, settings } : { ok: false, code: checked.code };
}
