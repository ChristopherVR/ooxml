import {
	DEFAULT_ANIMATION_DURATION_MS,
	DEFAULT_ANIMATION_TRIGGER,
	type PptxAnimationTimelineAnchor,
	type PptxElementAnimation,
} from 'ooxml-core/pptx';

import { setDelay, setDuration, setTrigger } from './animation-authoring';
import { editableAnimation } from './animation-painter';
import { buildAnimationTimelineRows, moveAnimationTimelineRowBy } from './animation-timeline-rows';

export interface AnimationEditingSlide {
	id?: string | undefined;
	animations?: readonly PptxElementAnimation[] | undefined;
	animationTimelineAnchors?: readonly PptxAnimationTimelineAnchor[] | undefined;
}

export type CommitRibbonAnimations = (animations: PptxElementAnimation[]) => void;

export interface RibbonAnimationEditing {
	slide: AnimationEditingSlide | undefined;
	elementId: string | undefined;
	commit: CommitRibbonAnimations | undefined;
}

export const RIBBON_ANIMATION_START_MODES = ['onClick', 'withPrevious', 'afterPrevious'] as const;
type StartMode = (typeof RIBBON_ANIMATION_START_MODES)[number];

export type RibbonAnimationTimingEdit =
	| { field: 'start'; value: StartMode }
	| { field: 'duration' | 'delay'; value: number }
	| { field: 'move'; value: -1 | 1 };

export function isRibbonAnimationStart(value: string): value is StartMode {
	return RIBBON_ANIMATION_START_MODES.some((mode) => mode === value);
}

export function ribbonAnimationTiming(editing: RibbonAnimationEditing | undefined) {
	const animations = editing?.slide?.animations ?? [];
	const anchors = editing?.slide?.animationTimelineAnchors ?? [];
	const entry = editableAnimation(animations, editing?.elementId);
	if (!entry || !editing?.commit) {
		return undefined;
	}
	const rows = buildAnimationTimelineRows(animations, anchors);
	const index = rows.findIndex((row) => row.kind === 'editor' && row.elementId === entry.elementId);
	const start = entry.trigger ?? DEFAULT_ANIMATION_TRIGGER;
	return {
		start: start === 'afterDelay' ? 'afterPrevious' : start,
		durationMs: entry.durationMs ?? DEFAULT_ANIMATION_DURATION_MS,
		delayMs: entry.delayMs ?? 0,
		earlier: index > 0,
		later: index >= 0 && index + 1 < rows.length,
	};
}

/** Route ribbon edits through the same operations as the Animation Pane. */
export function editRibbonAnimationTiming(
	editing: RibbonAnimationEditing | undefined,
	edit: RibbonAnimationTimingEdit,
): void {
	const timing = ribbonAnimationTiming(editing);
	const animations = editing?.slide?.animations;
	const elementId = editing?.elementId;
	const commit = editing?.commit;
	if (!timing || !animations || !elementId || !commit) {
		return;
	}
	switch (edit.field) {
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
