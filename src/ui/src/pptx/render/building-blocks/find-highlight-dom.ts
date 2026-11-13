/**
 * The canvas outline of the host find bar's matches, shared by every building-blocks canvas that
 * has no run-level highlight of its own (vanilla, Vue, Svelte, Angular): the elements holding a
 * match get `data-pptx-find-match` (`current` for the one with the current match), styled by each
 * binding's stylesheet. DOM only; call it again after every repaint of the slide.
 *
 * @module render/building-blocks/find-highlight-dom
 */
import type { FindResult } from '../find-replace';
import { findHighlightTargets } from './find-layout';
import type { FindHighlightTarget } from './find-layout';

/** The attribute set on an element holding a find match (`current` or `match`). */
export const FIND_MATCH_ATTRIBUTE = 'data-pptx-find-match';

/**
 * Mark exactly `targets` among the slide elements (`[data-element-id]`) under `root`; elements
 * that are no longer targets lose the mark. An empty list clears every mark.
 */
export function markFindHighlights(
	root: ParentNode,
	targets: readonly FindHighlightTarget[],
): void {
	const marks = new Map(
		targets.map((target) => [target.elementId, target.current ? 'current' : 'match']),
	);
	for (const node of root.querySelectorAll<HTMLElement>(`[${FIND_MATCH_ATTRIBUTE}]`)) {
		if (!marks.has(node.getAttribute('data-element-id') ?? '')) {
			node.removeAttribute(FIND_MATCH_ATTRIBUTE);
		}
	}
	if (marks.size === 0) {
		return;
	}
	for (const node of root.querySelectorAll<HTMLElement>('[data-element-id]')) {
		const mark = marks.get(node.getAttribute('data-element-id') ?? '');
		if (mark && node.getAttribute(FIND_MATCH_ATTRIBUTE) !== mark) {
			node.setAttribute(FIND_MATCH_ATTRIBUTE, mark);
		}
	}
}

/** Outline the matches of slide `slideIndex` (see {@link findHighlightTargets}). */
export function syncFindHighlights(
	root: ParentNode,
	matches: readonly FindResult[],
	currentIndex: number,
	slideIndex: number,
): void {
	markFindHighlights(root, findHighlightTargets(matches, currentIndex, slideIndex));
}
