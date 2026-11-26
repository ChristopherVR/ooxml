import {
	animationFor,
	captureNativeAnimationTemplate,
	DEFAULT_ANIMATION_TRIGGER,
	ooxmlToPresetName,
	type PptxElementAnimation,
	type PptxNativeAnimation,
} from 'ooxml-core/pptx';

import { directionForAnimationSubtype } from './animation-direction-options';
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

function importedPaintEffect(
	slide: AnimationEditingSlide,
	elementId: string | undefined,
	nodeId: number | undefined,
): PptxNativeAnimation | undefined {
	return slide.nativeAnimations?.find(
		(item) =>
			item.targetId &&
			item.presetClass &&
			item.presetId !== undefined &&
			(item.presetClass !== 'path' || item.motionPath) &&
			(nodeId === undefined ? item.targetId === elementId : item.nodeId === nodeId),
	);
}

export function canCaptureAnimationPaint(
	slide: AnimationEditingSlide,
	elementId: string | undefined,
	nodeId: number | undefined,
): boolean {
	return Boolean(
		(nodeId === undefined && editableAnimation(slide.animations ?? [], elementId)) ||
		importedPaintEffect(slide, elementId, nodeId),
	);
}

/** Capture a value snapshot and discard source-bound interactive trigger data. */
export function captureAnimationPaint(
	slide: AnimationEditingSlide,
	elementId: string | undefined,
	nodeId?: number,
): AnimationPaint | undefined {
	const entry =
		(nodeId === undefined ? editableAnimation(slide.animations ?? [], elementId) : undefined) ??
		importedPaintEntry(slide, importedPaintEffect(slide, elementId, nodeId));
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

function importedPaintEntry(
	slide: AnimationEditingSlide,
	effect: PptxNativeAnimation | undefined,
): PptxElementAnimation | undefined {
	const { targetId, presetClass, presetId } = effect ?? {};
	if (!effect || !targetId || !presetClass || presetId === undefined) return undefined;
	const slot = { entr: 'entrance', emph: 'emphasis', exit: 'exit', path: 'motionPath' }[
		presetClass
	];
	const token = ooxmlToPresetName({ presetClass, presetId }) ?? `${presetClass}.${presetId}`;
	const value = presetClass === 'path' ? effect.motionPath : token;
	if (!value) return undefined;
	return {
		elementId: targetId,
		[slot]: value,
		nativePaintTemplate:
			effect.nodeId === undefined || !slide.rawTiming
				? undefined
				: captureNativeAnimationTemplate(slide.rawTiming, effect.nodeId),
		durationMs: effect.durationMs,
		delayMs: effect.delayMs,
		trigger: effect.interactiveSequence ? effect.trigger : (effect.startMode ?? effect.trigger),
		direction: directionForAnimationSubtype(token, effect.presetSubtype),
		soundRId: effect.soundRId,
		soundPath: effect.soundPath,
		soundName: effect.soundName,
		stopSound: effect.stopSound,
	};
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
		// The saver copies the embedded part and creates a relationship for the destination slide.
		delete effect.soundRId;
		if (!effect.soundPath) delete effect.soundName;
	}
	const next = { ...effect, elementId, order };
	return existing
		? animations.map((entry) => (entry === existing ? next : entry))
		: [...animations, next];
}
