import { findHighlightTargets } from 'ooxml-ui/pptx';
import type { FindResult } from 'ooxml-ui/pptx';

/** The attribute the canvas block sets on an element holding a find match (`current` or `match`). */
export const FIND_MATCH_ATTRIBUTE = 'data-pptx-find-match';

/**
 * Outline the elements of the shown slide that hold the host find bar's matches, the current
 * one marked `current`. Elements that no longer match lose the mark.
 */
export function syncFindHighlights(
	root: ParentNode,
	matches: readonly FindResult[],
	currentIndex: number,
	slideIndex: number,
): void {
	const targets = new Map(
		findHighlightTargets(matches, currentIndex, slideIndex).map((target) => [
			target.elementId,
			target.current ? 'current' : 'match',
		]),
	);
	for (const node of root.querySelectorAll<HTMLElement>(`[${FIND_MATCH_ATTRIBUTE}]`)) {
		if (!targets.has(node.dataset.elementId ?? '')) {
			node.removeAttribute(FIND_MATCH_ATTRIBUTE);
		}
	}
	for (const node of root.querySelectorAll<HTMLElement>('[data-pptx-viewport] [data-element-id]')) {
		const mark = targets.get(node.dataset.elementId ?? '');
		if (mark && node.getAttribute(FIND_MATCH_ATTRIBUTE) !== mark) {
			node.setAttribute(FIND_MATCH_ATTRIBUTE, mark);
		}
	}
}
