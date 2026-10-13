import type { ElementAnimationState } from './animation-timeline-group';
import { BACKGROUND_ANIMATION_ID_SUFFIX } from './animation-target-id';

/** The active fixed value replaces authored opacity and leaves CSS ramps independent. */
export function getAnimationOpacity(
	state: ElementAnimationState | undefined,
	authored: number | string | undefined,
): number | string | undefined {
	return state?.opacity ?? authored;
}

const authoredOpacity = new WeakMap<HTMLElement, string>();

/** Apply/restore a fixed opacity without erasing an element's authored alpha. */
export function applyAnimationOpacity(
	element: HTMLElement,
	state: ElementAnimationState | undefined,
): void {
	if (state?.opacity !== undefined) {
		if (!authoredOpacity.has(element)) {
			authoredOpacity.set(
				element,
				element.dataset.animId?.endsWith(BACKGROUND_ANIMATION_ID_SUFFIX)
					? ''
					: element.style.opacity,
			);
		}
		element.style.opacity = String(getAnimationOpacity(state, undefined));
	} else if (authoredOpacity.has(element)) {
		element.style.opacity = authoredOpacity.get(element)!;
		authoredOpacity.delete(element);
	}
}
