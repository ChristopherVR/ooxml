import {
	animationFor,
	DEFAULT_ANIMATION_TRIGGER,
	type PptxElementAnimation,
} from 'ooxml-core/pptx';

import type { AnimationEditingSlide } from './ribbon-animation-editing';

export function editableAnimation(
	animations: readonly PptxElementAnimation[],
	elementId: string | undefined,
): PptxElementAnimation | undefined {
	const entry = elementId === undefined ? undefined : animationFor(animations, elementId);
	return entry &&
		(entry.motionPath ||
			[entry.entrance, entry.emphasis, entry.exit].some((preset) => preset && preset !== 'none'))
		? entry
		: undefined;
}

export interface AnimationPaint {
	source: { slideId: string | undefined; elementId: string };
	effect: Omit<PptxElementAnimation, 'elementId' | 'order'>;
}

/** Capture a value snapshot and discard source-bound interactive trigger data. */
export function captureAnimationPaint(
	slide: AnimationEditingSlide,
	elementId: string | undefined,
): AnimationPaint | undefined {
	const entry = editableAnimation(slide.animations ?? [], elementId);
	if (!entry) {
		return undefined;
	}
	const { elementId: sourceElementId, ...effect } = entry;
	// Template XML can contain the source shape and time node IDs.
	delete effect.order;
	delete effect.orderFromTimeline;
	delete effect.triggerShapeId;
	delete effect.triggerBookmark;
	delete effect.buildTemplates;
	delete effect.nativeAfterAnimation;
	delete effect.nativeAfterAnimationColor;
	if (entry.trigger && ['onShapeClick', 'onMediaBookmark', 'onHover'].includes(entry.trigger)) {
		effect.trigger = DEFAULT_ANIMATION_TRIGGER;
	}
	return { source: { slideId: slide.id, elementId: sourceElementId }, effect };
}

/** Replace the target's editor effects in place, or append after the full timeline. */
export function applyAnimationPaint(
	paint: AnimationPaint,
	slide: AnimationEditingSlide,
	elementId: string,
): PptxElementAnimation[] | undefined {
	const sameSlide = paint.source.slideId === slide.id;
	if (sameSlide && paint.source.elementId === elementId) {
		return undefined;
	}
	const animations = slide.animations ?? [];
	const existing = animationFor(animations, elementId);
	const order = existing
		? existing.order
		: Math.max(
				-1,
				...animations.map((entry) => entry.order ?? 0),
				...(slide.animationTimelineAnchors ?? []).map((anchor) => anchor.order),
			) + 1;
	const effect = { ...paint.effect };
	if (!sameSlide && effect.soundRId) {
		// Embedded sound relationships belong to a single slide part.
		delete effect.soundRId;
		delete effect.soundPath;
		delete effect.soundName;
	}
	const next = { ...effect, elementId, order };
	return existing
		? animations.map((entry) => (entry === existing ? next : entry))
		: [...animations, next];
}
