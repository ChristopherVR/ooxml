/**
 * `animation-after-effect-authoring`  -  pure decision logic for the animation
 * panel's "after animation" control (dim to colour / hide after animation /
 * hide on next click / don't dim).
 *
 * Playback of `afterAnimation` / `afterAnimationColor` was already fully
 * implemented (see `animation-after-effect.ts`'s
 * `applyAfterAnimationFromEditorList`); what was missing was an authoring
 * control to set them. This module is the write side's counterpart: value
 * catalog + immutable setters, matching the shape of `animation-authoring.ts`
 * (which already owns `entrance` / `exit` / `emphasis` / timing setters for
 * the same `PptxElementAnimation` entry).
 *
 * @module services/editor-animation-after-effect
 */
import type { PptxAfterAnimationAction, PptxElementAnimation, PptxNativeAnimation } from '../types';

import { animationFor, upsert } from './editor-animation-entry';
import { assertAfterAnimationClearSupported } from './animation-after-effect-write';
import { indexEffectNodes } from './animation-timing-tree';
import type { XmlObject } from '../types';

/** Option values for the "after animation" selector, in the order PowerPoint lists them. */
export const AFTER_ANIMATION_VALUES: readonly PptxAfterAnimationAction[] = [
	'none',
	'dimToColor',
	'hideAfterAnimation',
	'hideOnNextClick',
];

/** Default dim colour offered the first time a user picks "Dim after animation". */
export const DEFAULT_AFTER_ANIMATION_DIM_COLOR = '#808080';

/**
 * Sets (or clears) the "after animation" action for the element's animation
 * entry. Switching AWAY from `dimToColor` clears `afterAnimationColor`, so a
 * value left over from a prior dim choice can't silently reactivate if the
 * user picks `dimToColor` again later; switching back in supplies a fresh
 * default colour instead. Callers that want to restore a specific colour
 * should call `setAfterAnimationColor` afterwards.
 */
export function setAfterAnimation(
	anims: readonly PptxElementAnimation[],
	elementId: string,
	action: PptxAfterAnimationAction | undefined,
): PptxElementAnimation[] {
	return upsert(anims, elementId, (cur) => {
		const value = action;
		return {
			...cur,
			afterAnimation: value,
			afterAnimationColor:
				value === 'dimToColor'
					? (cur.afterAnimationColor ?? DEFAULT_AFTER_ANIMATION_DIM_COLOR)
					: undefined,
		};
	});
}

/**
 * Sets the dim-to colour for an entry already using `afterAnimation ===
 * 'dimToColor'`. A no-op (returns a shallow copy) when the entry does not
 * have `dimToColor` active, so a colour swatch left mounted during an
 * unrelated re-render can't accidentally turn dimming on.
 */
export function setAfterAnimationColor(
	anims: readonly PptxElementAnimation[],
	elementId: string,
	color: string,
): PptxElementAnimation[] {
	const entry = animationFor(anims, elementId);
	if (!entry || (entry.afterAnimation ?? entry.nativeAfterAnimation) !== 'dimToColor') {
		return [...anims];
	}
	return upsert(anims, elementId, (cur) => ({
		...cur,
		afterAnimation: 'dimToColor',
		afterAnimationColor: color,
	}));
}

/** Reads the current "after animation" action, defaulting to `'none'`. */
export function getAfterAnimation(
	slideAnimations: readonly PptxElementAnimation[],
	elementId: string,
): PptxAfterAnimationAction {
	const entry = animationFor(slideAnimations, elementId);
	return entry?.afterAnimation ?? entry?.nativeAfterAnimation ?? 'none';
}

export function getAfterAnimationState(entry: PptxElementAnimation | undefined): {
	action: PptxAfterAnimationAction;
	color: string | undefined;
	showColor: boolean;
} {
	const action = entry?.afterAnimation ?? entry?.nativeAfterAnimation ?? 'none';
	return {
		action,
		color:
			action === 'dimToColor'
				? (entry?.afterAnimationColor ?? entry?.nativeAfterAnimationColor)
				: undefined,
		showColor:
			entry?.afterAnimation !== undefined ||
			/^#[0-9a-f]{6}$/iu.test(entry?.nativeAfterAnimationColor ?? ''),
	};
}

export function isNativeAfterAnimationOnly(entry: PptxElementAnimation | undefined): boolean {
	return Boolean(
		entry?.nativeAfterAnimation &&
		!entry.entrance &&
		!entry.exit &&
		!entry.emphasis &&
		!entry.motionPath,
	);
}

/** Expose imported whole-shape after-state without creating an authored override. */
export function mergeNativeAfterAnimationIntoEditorAnimations(
	nativeAnimations: readonly PptxNativeAnimation[] | undefined,
	animations: PptxElementAnimation[],
	rawTiming?: XmlObject,
): void {
	for (const native of nativeAnimations ?? []) {
		const existing = native.targetId ? animationFor(animations, native.targetId) : undefined;
		const action =
			native.afterAnimationAction ?? (existing?.afterAnimation !== undefined ? 'none' : undefined);
		if (
			(nativeAnimations ?? []).filter(
				(other) =>
					other.targetId === native.targetId && ['entr', 'emph'].includes(other.presetClass ?? ''),
			).length !== 1
		)
			continue;
		if (
			!native.targetId ||
			!action ||
			!['entr', 'emph'].includes(native.presetClass ?? '') ||
			native.textTarget
		)
			continue;
		if (
			native.target &&
			(native.target.type !== 'shape' ||
				native.target.backgroundOnly ||
				native.target.subShapeId ||
				native.target.graphicElement ||
				native.target.oleChartElement)
		)
			continue;
		try {
			if (
				rawTiming &&
				native.target?.type === 'shape' &&
				indexEffectNodes(rawTiming).filter(
					(ref) =>
						ref.spid === (native.target?.type === 'shape' ? native.target.shapeId : undefined) &&
						['entr', 'emph'].includes(ref.presetClass),
				).length !== 1
			)
				continue;
			const effect = rawTiming
				? indexEffectNodes(rawTiming).find(
						(ref) =>
							String(ref.cTn['@_id']) === String(native.nodeId) &&
							native.target?.type === 'shape' &&
							ref.spid === native.target.shapeId,
					)
				: undefined;
			if (rawTiming && !effect) continue;
			assertAfterAnimationClearSupported(
				effect?.cTn ?? ({ 'p:subTnLst': native.cTnAttributes?.['p:subTnLst'] } as XmlObject),
			);
		} catch {
			// Imported unmodeled after-state remains raw and gets no editing control.
			continue;
		}
		let entry = animationFor(animations, native.targetId);
		if (!entry) {
			entry = { elementId: native.targetId };
			animations.push(entry);
		}
		if (entry.nativeAfterAnimation !== undefined) continue;
		entry.nativeAfterAnimation = action;
		entry.nativeAfterAnimationColor = native.afterAnimationColor;
	}
}

export function isNativeAfterAnimationFallback(entry: PptxElementAnimation): boolean {
	return (
		entry.nativeAfterAnimation !== undefined &&
		Object.entries(entry).every(
			([key, value]) =>
				value === undefined ||
				[
					'elementId',
					'nativeAfterAnimation',
					'nativeAfterAnimationColor',
					'soundRId',
					'soundPath',
					'soundName',
					'stopSound',
				].includes(key),
		)
	);
}
