/**
 * Slide jumps on text runs that name their slide by index only.
 *
 * A run's `a:hlinkClick action="ppaction://hlinksldjump"` must point at a
 * `…/relationships/slide` relationship. A loaded run keeps the relationship it
 * was parsed with (`hyperlinkRId`, `hyperlink` = its target). A run made in
 * code gives only `hyperlinkTargetSlideIndex`; it has no relationship yet, so
 * the save resolves the index to the slide part, the way shape-level actions
 * do (`serializeSingleAction`), and the run writer then adds the relationship
 * for that target. An index that names no slide of the saved deck is reported
 * through `onMissing` and the link is not written: PowerPoint does not open a
 * file with a slide jump that has no slide relationship.
 *
 * @module runtime/run-slide-jump-targets
 */

import type { TextSegment, TextStyle } from '../../types';

export const SLIDE_JUMP_ACTION = 'ppaction://hlinksldjump';

/** The `Target`, from the part being saved, of the deck's slide at `index`. */
export type SlideTargetAt = (index: number) => string | undefined;

/** Called for a slide jump whose index names no slide of the deck being saved. */
export type MissingSlideTarget = (index: number) => void;

/**
 * `style` with its slide jump pointing at the slide part, or `style` itself
 * when it has no index-only slide jump (every loaded run, every URL link).
 */
export function withRunSlideJumpTarget<T extends TextStyle | undefined>(
	style: T,
	slideTargetAt: SlideTargetAt | undefined,
	onMissing?: MissingSlideTarget,
): T {
	if (!style || !slideTargetAt) {
		return style;
	}
	const index = style.hyperlinkTargetSlideIndex;
	if (typeof index !== 'number' || style.hyperlinkRId) {
		return style;
	}
	const hyperlink = style.hyperlink?.trim();
	if (hyperlink && !hyperlink.startsWith('ppaction://')) {
		return style;
	}
	if (style.hyperlinkAction && style.hyperlinkAction !== SLIDE_JUMP_ACTION) {
		return style;
	}
	const target = slideTargetAt(index);
	if (!target) {
		// A slide jump without its slide relationship makes PowerPoint refuse
		// the file, so the link is left out of the run and reported.
		onMissing?.(index);
		const {
			hyperlink: _hyperlink,
			hyperlinkAction: _action,
			hyperlinkTargetSlideIndex: _index,
			hyperlinkTooltip: _tooltip,
			...rest
		} = style;
		return rest as T;
	}
	return { ...style, hyperlink: target, hyperlinkAction: SLIDE_JUMP_ACTION };
}

/** {@link withRunSlideJumpTarget} for every segment; the same array when none changes. */
export function withSegmentSlideJumpTargets(
	segments: TextSegment[] | undefined,
	slideTargetAt: SlideTargetAt | undefined,
	onMissing?: MissingSlideTarget,
): TextSegment[] | undefined {
	if (!segments || !slideTargetAt) {
		return segments;
	}
	let changed = false;
	const next = segments.map((segment) => {
		const style = withRunSlideJumpTarget(segment.style, slideTargetAt, onMissing);
		if (style === segment.style) {
			return segment;
		}
		changed = true;
		return { ...segment, style };
	});
	return changed ? next : segments;
}
