import {
	DEFAULT_ANIMATION_DURATION_MS,
	DEFAULT_ANIMATION_TRIGGER,
	editNativeEffectTiming,
	nativeAnimationStartAvailability,
	type PptxAnimationTimelineAnchor,
	type PptxAnimationDirection,
	type PptxElementAnimation,
	type PptxElement,
	type PptxNativeAnimation,
	type PptxSlide,
} from 'ooxml-core/pptx';

import { setDelay, setDirection, setDuration, setTrigger } from './animation-authoring';
import { editableAnimation } from './animation-painter';
import { buildAnimationTimelineRows, moveAnimationTimelineRowBy } from './animation-timeline-rows';

export interface AnimationEditingSlide {
	id?: string | undefined;
	elements?: readonly PptxElement[] | undefined;
	animations?: readonly PptxElementAnimation[] | undefined;
	animationTimelineAnchors?: readonly PptxAnimationTimelineAnchor[] | undefined;
	nativeAnimations?: readonly PptxNativeAnimation[] | undefined;
	rawTiming?: PptxSlide['rawTiming'];
}

export type CommitRibbonAnimations = (
	animations: PptxElementAnimation[],
	timing?: Pick<PptxSlide, 'rawTiming' | 'nativeAnimations'>,
) => void;

export interface RibbonAnimationEditing {
	slide: AnimationEditingSlide | undefined;
	elementId: string | undefined;
	commit: CommitRibbonAnimations | undefined;
	nativeEffectId?: number | undefined;
}

export const RIBBON_ANIMATION_START_MODES = ['onClick', 'withPrevious', 'afterPrevious'] as const;
type StartMode = (typeof RIBBON_ANIMATION_START_MODES)[number];

export type RibbonAnimationTimingEdit =
	| { field: 'start'; value: StartMode }
	| { field: 'duration' | 'delay'; value: number }
	| { field: 'direction'; value: PptxAnimationDirection }
	| { field: 'move'; value: -1 | 1 };

export function isRibbonAnimationStart(value: string): value is StartMode {
	return RIBBON_ANIMATION_START_MODES.some((mode) => mode === value);
}

export function nativeRibbonTimingTarget(
	editing: RibbonAnimationEditing | undefined,
): PptxNativeAnimation | undefined {
	const slide = editing?.slide;
	if (!editing || !slide?.rawTiming) return undefined;
	return slide.nativeAnimations?.find(
		(effect) =>
			effect.nodeId !== undefined &&
			effect.targetId &&
			effect.presetClass &&
			(editing.nativeEffectId !== undefined
				? effect.nodeId === editing.nativeEffectId
				: effect.targetId === editing.elementId),
	);
}

export function ribbonAnimationTiming(editing: RibbonAnimationEditing | undefined) {
	const animations = editing?.slide?.animations ?? [];
	const anchors = editing?.slide?.animationTimelineAnchors ?? [];
	const entry = editableAnimation(animations, editing?.elementId);
	if (!editing?.commit) {
		return undefined;
	}
	const native =
		!entry || editing.nativeEffectId !== undefined ? nativeRibbonTimingTarget(editing) : undefined;
	if (editing.nativeEffectId !== undefined && !native) return undefined;
	if (native?.nodeId !== undefined && editing.slide?.rawTiming && editing.slide.nativeAnimations) {
		const availability = nativeAnimationStartAvailability(
			editing.slide.rawTiming,
			editing.slide.nativeAnimations,
			native.nodeId,
		);
		const start =
			(native.interactiveSequence ? native.trigger : (native.startMode ?? native.trigger)) ??
			DEFAULT_ANIMATION_TRIGGER;
		return {
			start: start === 'afterDelay' ? 'afterPrevious' : start,
			durationMs: native.durationMs ?? DEFAULT_ANIMATION_DURATION_MS,
			delayMs: native.delayMs ?? 0,
			earlier: false,
			later: false,
			startOptions: availability.options,
			startRefusal: availability.refusal,
		};
	}
	if (!entry) return undefined;
	const rows = buildAnimationTimelineRows(animations, anchors);
	const index = rows.findIndex((row) => row.kind === 'editor' && row.elementId === entry.elementId);
	const start = entry.trigger ?? DEFAULT_ANIMATION_TRIGGER;
	return {
		start: start === 'afterDelay' ? 'afterPrevious' : start,
		durationMs: entry.durationMs ?? DEFAULT_ANIMATION_DURATION_MS,
		delayMs: entry.delayMs ?? 0,
		earlier: index > 0,
		later: index >= 0 && index + 1 < rows.length,
		startOptions: RIBBON_ANIMATION_START_MODES,
		startRefusal: undefined,
	};
}

/** Route ribbon edits through the same operations as the Animation Pane. */
export function editRibbonAnimationTiming(
	editing: RibbonAnimationEditing | undefined,
	edit: RibbonAnimationTimingEdit,
): 'stale' | void {
	if (editing?.nativeEffectId !== undefined && !nativeRibbonTimingTarget(editing)) return 'stale';
	const timing = ribbonAnimationTiming(editing);
	const slide = editing?.slide;
	const animations = slide?.animations ?? [];
	const elementId = editing?.elementId;
	const commit = editing?.commit;
	if (!timing || !commit) {
		return;
	}
	const native =
		editing?.nativeEffectId !== undefined || !editableAnimation(animations, elementId)
			? nativeRibbonTimingTarget(editing)
			: undefined;
	if (
		native?.nodeId !== undefined &&
		slide?.rawTiming &&
		slide.nativeAnimations &&
		edit.field !== 'move'
	) {
		const patch = editNativeEffectTiming(
			slide.rawTiming,
			slide.nativeAnimations,
			native.nodeId,
			edit,
		);
		if (patch) commit([...animations], patch);
		return;
	}
	if (!elementId) return;
	switch (edit.field) {
		case 'direction':
			commit(setDirection(animations, elementId, edit.value));
			break;
		case 'start':
			if (isRibbonAnimationStart(edit.value) && timing.start !== edit.value) {
				commit(setTrigger(animations, elementId, edit.value));
			}
			break;
		case 'duration':
		case 'delay': {
			if (!Number.isFinite(edit.value)) {
				break;
			}
			const setter = edit.field === 'duration' ? setDuration : setDelay;
			const field = edit.field === 'duration' ? 'durationMs' : 'delayMs';
			const updated = setter(animations, elementId, Math.round(edit.value));
			if (editableAnimation(updated, elementId)?.[field] !== timing[field]) {
				commit(updated);
			}
			break;
		}
		case 'move':
			if ((edit.value === -1 && timing.earlier) || (edit.value === 1 && timing.later)) {
				commit(
					moveAnimationTimelineRowBy(
						animations,
						editing.slide?.animationTimelineAnchors ?? [],
						elementId,
						edit.value,
					),
				);
			}
			break;
	}
}
