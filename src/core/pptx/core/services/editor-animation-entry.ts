import type { PptxElementAnimation } from '../types';

export const DEFAULT_ANIMATION_DURATION_MS = 500;
export const DEFAULT_ANIMATION_TRIGGER = 'onClick' as const;

/** Literal none presets do not create an effect or justify an after-state target. */
export function hasAuthoredAnimationEffect(entry: PptxElementAnimation): boolean {
	return Boolean(
		[entry.entrance, entry.exit, entry.emphasis].some((preset) => preset && preset !== 'none') ||
		entry.motionPath,
	);
}

export function hasAnimation(
	animations: readonly PptxElementAnimation[],
	elementId: string,
): boolean {
	const entry = animationFor(animations, elementId);
	return Boolean(entry && (entry.entrance || entry.exit || entry.emphasis || entry.motionPath));
}

export function animationFor(
	slideAnimations: readonly PptxElementAnimation[],
	elementId: string,
): PptxElementAnimation | undefined {
	return slideAnimations.find((a) => a.elementId === elementId);
}

export function upsert(
	anims: readonly PptxElementAnimation[],
	elementId: string,
	updater: (current: PptxElementAnimation) => PptxElementAnimation | null,
): PptxElementAnimation[] {
	const idx = anims.findIndex((a) => a.elementId === elementId);
	if (idx >= 0) {
		const updated = updater({ ...anims[idx] });
		if (updated === null) {
			return anims.filter((a) => a.elementId !== elementId);
		}
		return anims.map((a, i) => (i === idx ? updated : a));
	}
	// No existing entry  -  create one and then apply the updater.
	const created: PptxElementAnimation = {
		elementId,
		durationMs: DEFAULT_ANIMATION_DURATION_MS,
		order: anims.length,
		trigger: DEFAULT_ANIMATION_TRIGGER,
	};
	const updated = updater(created);
	if (updated === null) {
		return [...anims];
	}
	return [...anims, updated];
}
