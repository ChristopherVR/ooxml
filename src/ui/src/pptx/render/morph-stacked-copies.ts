/**
 * Incoming elements that are near-exact copies stacked on one another.
 *
 * PowerPoint dissolves a slide's new objects in as one layer, so two copies
 * of the same glyph a pixel or two apart read as one glyph the whole way in.
 * Each binding fades every element on its own, and two half-transparent
 * copies show both of their edges: issue #141's slide 7 stacks two "L" text
 * boxes 2px apart, and the morph from slide 6 drew a doubled "L" until the
 * fade finished. Only the topmost copy fades; the ones beneath hold until the
 * last frame, where the opaque copy above already covers them.
 *
 * @module render/morph-stacked-copies
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { mediaIdentity } from './morph-predicates';

/** How far apart (px) two copies may sit and still read as one object. */
const MAX_COPY_OFFSET = 4;

function ownText(element: PptxElement): string {
	return ((element as { text?: string }).text ?? '').replace(/\s+/gu, ' ').trim();
}

/** Whether `a` and `b` paint the same thing in (almost) the same place. */
function isStackedCopy(a: PptxElement, b: PptxElement): boolean {
	if (a.type !== b.type || (a.rotation ?? 0) !== (b.rotation ?? 0)) {
		return false;
	}
	if (Math.abs(a.width - b.width) > 1 || Math.abs(a.height - b.height) > 1) {
		return false;
	}
	if (Math.hypot(a.x - b.x, a.y - b.y) > MAX_COPY_OFFSET) {
		return false;
	}
	const text = ownText(a);
	if (text !== '') {
		return text === ownText(b);
	}
	const media = mediaIdentity(a);
	return media !== undefined && media === mediaIdentity(b);
}

/**
 * Ids of the incoming elements that a later (higher) element in `elements`
 * copies. `elements` is in paint order, bottom first.
 */
export function coveredStackedCopyIds(elements: readonly PptxElement[]): Set<string> {
	const covered = new Set<string>();
	for (let i = 0; i < elements.length; i++) {
		const lower = elements[i] as PptxElement;
		for (let j = i + 1; j < elements.length; j++) {
			if (isStackedCopy(lower, elements[j] as PptxElement)) {
				covered.add(lower.id);
				break;
			}
		}
	}
	return covered;
}
